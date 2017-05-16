VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0100C 
   Caption         =   "客户特殊要求材质输入 - AQA0100C"
   ClientHeight    =   9315
   ClientLeft      =   135
   ClientTop       =   1995
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
   ScaleHeight     =   9315
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin InDate.ULabel ulb_Detail 
      Height          =   310
      Left            =   6300
      Top             =   150
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   0
      BackColor       =   16777215
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
   Begin VB.TextBox txt_CUST_SPEC_NO 
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
      Left            =   1890
      MaxLength       =   9
      TabIndex        =   0
      Top             =   150
      Width           =   2565
   End
   Begin VB.TextBox txt_SMP_LOC 
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
      Left            =   1890
      MaxLength       =   1
      TabIndex        =   1
      Top             =   555
      Width           =   555
   End
   Begin VB.TextBox txt_SMP_LOC_NAME 
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
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   5
      Top             =   555
      Width           =   1995
   End
   Begin VB.TextBox txt_INS_EMP 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10410
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   405
   End
   Begin CSTextLibCtl.sidbEdit txt_SMP_LEN 
      Height          =   310
      Left            =   6300
      TabIndex        =   2
      Top             =   555
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   547
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit txt_PRE_SMP_QTY 
      Height          =   310
      Left            =   9420
      TabIndex        =   3
      Top             =   555
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   547
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   150
      Top             =   150
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "客户特殊要求编号"
      Alignment       =   0
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   1
      Left            =   4560
      Top             =   150
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "说明"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   150
      Top             =   555
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "取试料位置"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   1
      Left            =   4560
      Top             =   555
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "取试料长度"
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
      Index           =   0
      Left            =   7680
      Top             =   555
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "复样数量"
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
   Begin TabDlg.SSTab ssT 
      Height          =   8130
      Left            =   75
      TabIndex        =   6
      Top             =   975
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14340
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   14737632
      TabCaption(0)   =   "拉伸试验"
      TabPicture(0)   =   "AQA0100C.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt_RA_DIR_NAME"
      Tab(0).Control(1)=   "txt_RA_DIR_CD"
      Tab(0).Control(2)=   "txt_YR_DSC_CD"
      Tab(0).Control(3)=   "txt_SP_EL_CD(0)"
      Tab(0).Control(4)=   "txt_EL_CD(0)"
      Tab(0).Control(5)=   "txt_EL_CD(1)"
      Tab(0).Control(6)=   "txt_SP_EL_CD(1)"
      Tab(0).Control(7)=   "txt_SG_EL_CD(0)"
      Tab(0).Control(8)=   "txt_SG_EL_CD(1)"
      Tab(0).Control(9)=   "txt_SNPP_EL_CD(0)"
      Tab(0).Control(10)=   "txt_SNPP_EL_CD(1)"
      Tab(0).Control(11)=   "txt_SP_EL_SMP_CD"
      Tab(0).Control(12)=   "txt_TENCIL_SMP_CD"
      Tab(0).Control(13)=   "txt_YP_DSC_CD"
      Tab(0).Control(14)=   "txt_SP_EL_DSC_CD"
      Tab(0).Control(15)=   "txt_SG_EL_DSC_CD"
      Tab(0).Control(16)=   "txt_SNPP_EL_DSC_CD"
      Tab(0).Control(17)=   "txt_EL_DSC_CD"
      Tab(0).Control(18)=   "txt_RA_DSC_CD"
      Tab(0).Control(19)=   "txt_TS_DSC_CD"
      Tab(0).Control(20)=   "ULabel4(1)"
      Tab(0).Control(21)=   "ULabel4(0)"
      Tab(0).Control(22)=   "ULabel4(21)"
      Tab(0).Control(23)=   "ULabel4(20)"
      Tab(0).Control(24)=   "ULabel4(19)"
      Tab(0).Control(25)=   "ULabel4(18)"
      Tab(0).Control(26)=   "ULabel4(17)"
      Tab(0).Control(27)=   "ULabel4(16)"
      Tab(0).Control(28)=   "ULabel4(2)"
      Tab(0).Control(29)=   "ULabel4(3)"
      Tab(0).Control(30)=   "ULabel4(4)"
      Tab(0).Control(31)=   "ULabel4(5)"
      Tab(0).Control(32)=   "ULabel4(7)"
      Tab(0).Control(33)=   "ULabel4(6)"
      Tab(0).Control(34)=   "ULabel4(8)"
      Tab(0).Control(35)=   "ULabel4(22)"
      Tab(0).Control(36)=   "sdb_YP_MIN"
      Tab(0).Control(37)=   "sdb_YP_MAX"
      Tab(0).Control(38)=   "ULabel4(9)"
      Tab(0).Control(39)=   "sdb_TS_MIN"
      Tab(0).Control(40)=   "sdb_TS_MAX"
      Tab(0).Control(41)=   "ULabel4(10)"
      Tab(0).Control(42)=   "sdb_RA_MIN"
      Tab(0).Control(43)=   "sdb_RA_MAX"
      Tab(0).Control(44)=   "ULabel4(11)"
      Tab(0).Control(45)=   "sdb_EL_MIN"
      Tab(0).Control(46)=   "sdb_EL_MAX"
      Tab(0).Control(47)=   "ULabel4(12)"
      Tab(0).Control(48)=   "sdb_SNPP_EL_MIN"
      Tab(0).Control(49)=   "sdb_SNPP_EL_MAX"
      Tab(0).Control(50)=   "ULabel4(13)"
      Tab(0).Control(51)=   "sdb_SG_EL_MIN"
      Tab(0).Control(52)=   "sdb_SG_EL_MAX"
      Tab(0).Control(53)=   "ULabel4(14)"
      Tab(0).Control(54)=   "sdb_SP_EL_MIN"
      Tab(0).Control(55)=   "sdb_SP_EL_MAX"
      Tab(0).Control(56)=   "ULabel4(15)"
      Tab(0).Control(57)=   "ULabel4(59)"
      Tab(0).Control(58)=   "sdb_YR_MIN"
      Tab(0).Control(59)=   "sdb_YR_MAX"
      Tab(0).Control(60)=   "ULabel4(60)"
      Tab(0).Control(61)=   "Line3(0)"
      Tab(0).Control(62)=   "Line3(1)"
      Tab(0).Control(63)=   "Line3(2)"
      Tab(0).Control(64)=   "Line3(3)"
      Tab(0).Control(65)=   "Line3(4)"
      Tab(0).ControlCount=   66
      TabCaption(1)   =   "高温拉伸试验"
      TabPicture(1)   =   "AQA0100C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line3(9)"
      Tab(1).Control(1)=   "Line3(8)"
      Tab(1).Control(2)=   "Line3(7)"
      Tab(1).Control(3)=   "Line3(6)"
      Tab(1).Control(4)=   "Line3(5)"
      Tab(1).Control(5)=   "ULabel4(31)"
      Tab(1).Control(6)=   "ULabel4(30)"
      Tab(1).Control(7)=   "ULabel4(29)"
      Tab(1).Control(8)=   "ULabel4(28)"
      Tab(1).Control(9)=   "ULabel4(27)"
      Tab(1).Control(10)=   "ULabel4(26)"
      Tab(1).Control(11)=   "ULabel4(25)"
      Tab(1).Control(12)=   "ULabel4(24)"
      Tab(1).Control(13)=   "ULabel4(23)"
      Tab(1).Control(14)=   "ULabel27(18)"
      Tab(1).Control(15)=   "sdb_HGT_EL_MAX"
      Tab(1).Control(16)=   "sdb_HGT_EL_MIN"
      Tab(1).Control(17)=   "ULabel27(19)"
      Tab(1).Control(18)=   "sdb_HGT_SNPP_EL_MAX"
      Tab(1).Control(19)=   "sdb_HGT_SNPP_EL_MIN"
      Tab(1).Control(20)=   "ULabel27(20)"
      Tab(1).Control(21)=   "sdb_HGT_SP_EL_MAX"
      Tab(1).Control(22)=   "sdb_HGT_SP_EL_MIN"
      Tab(1).Control(23)=   "ULabel27(17)"
      Tab(1).Control(24)=   "sdb_HGT_RA_MAX"
      Tab(1).Control(25)=   "sdb_HGT_RA_MIN"
      Tab(1).Control(26)=   "ULabel27(16)"
      Tab(1).Control(27)=   "sdb_HGT_TS_MAX"
      Tab(1).Control(28)=   "sdb_HGT_TS_MIN"
      Tab(1).Control(29)=   "ULabel27(15)"
      Tab(1).Control(30)=   "sdb_HGT_YP_MAX"
      Tab(1).Control(31)=   "sdb_HGT_YP_MIN"
      Tab(1).Control(32)=   "ULabel27(2)"
      Tab(1).Control(33)=   "ULabel27(7)"
      Tab(1).Control(34)=   "ULabel27(6)"
      Tab(1).Control(35)=   "ULabel27(5)"
      Tab(1).Control(36)=   "ULabel27(4)"
      Tab(1).Control(37)=   "ULabel27(3)"
      Tab(1).Control(38)=   "txt_HGT_TENCIL_TMP_UNIT"
      Tab(1).Control(39)=   "txt_HGT_TENCIL_TMP"
      Tab(1).Control(40)=   "txt_HGT_TENCIL_SMP_CD"
      Tab(1).Control(41)=   "txt_HGT_SP_EL_SMP_CD"
      Tab(1).Control(42)=   "txt_HGT_EL_CD(1)"
      Tab(1).Control(43)=   "txt_HGT_EL_CD(0)"
      Tab(1).Control(44)=   "txt_HGT_SNPP_EL_CD(1)"
      Tab(1).Control(45)=   "txt_HGT_SNPP_EL_CD(0)"
      Tab(1).Control(46)=   "txt_HGT_SP_EL_CD(1)"
      Tab(1).Control(47)=   "txt_HGT_SP_EL_CD(0)"
      Tab(1).Control(48)=   "txt_HGT_YP_DSC_CD"
      Tab(1).Control(49)=   "txt_HGT_SP_EL_DSC_CD"
      Tab(1).Control(50)=   "txt_HGT_RA_DSC_CD"
      Tab(1).Control(51)=   "txt_HGT_TS_DSC_CD"
      Tab(1).Control(52)=   "txt_HGT_SNPP_EL_DSC_CD"
      Tab(1).Control(53)=   "txt_HGT_EL_DSC_CD"
      Tab(1).ControlCount=   54
      TabCaption(2)   =   "冲击、时效"
      TabPicture(2)   =   "AQA0100C.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_TIM_IMPACT_DIR(1)"
      Tab(2).Control(1)=   "txt_TIM_IMPACT_DIR(0)"
      Tab(2).Control(2)=   "txt_IMPACT_DIR(1)"
      Tab(2).Control(3)=   "txt_IMPACT_DIR(0)"
      Tab(2).Control(4)=   "txt_TIM_IMPACT_DSC_CD"
      Tab(2).Control(5)=   "txt_IMPACT_DSC_CD"
      Tab(2).Control(6)=   "txt_IMPACT_SMP_CD"
      Tab(2).Control(7)=   "txt_IMPACT_KND(0)"
      Tab(2).Control(8)=   "txt_IMPACT_KND(1)"
      Tab(2).Control(9)=   "txt_TIM_IMPACT_KND(1)"
      Tab(2).Control(10)=   "txt_TIM_IMPACT_TMP"
      Tab(2).Control(11)=   "txt_TIM_IMPACT_TMP_UNIT"
      Tab(2).Control(12)=   "txt_TIM_IMPACT_SMP_CD"
      Tab(2).Control(13)=   "txt_TIM_IMPACT_KND(0)"
      Tab(2).Control(14)=   "txt_IMPACT_TMP"
      Tab(2).Control(15)=   "txt_IMPACT_TMP_UNIT"
      Tab(2).Control(16)=   "txt_A_IMPACT_TMP_UNIT"
      Tab(2).Control(17)=   "txt_A_IMPACT_TMP"
      Tab(2).Control(18)=   "txt_A_IMPACT_KND(1)"
      Tab(2).Control(19)=   "txt_A_IMPACT_KND(0)"
      Tab(2).Control(20)=   "txt_A_IMPACT_SMP_CD"
      Tab(2).Control(21)=   "txt_A_IMPACT_DSC_CD"
      Tab(2).Control(22)=   "txt_A_IMPACT_DIR(0)"
      Tab(2).Control(23)=   "txt_A_IMPACT_DIR(1)"
      Tab(2).Control(24)=   "txt_A_TIM_IMPACT_KND(0)"
      Tab(2).Control(25)=   "txt_A_TIM_IMPACT_SMP_CD"
      Tab(2).Control(26)=   "txt_A_TIM_IMPACT_TMP_UNIT"
      Tab(2).Control(27)=   "txt_A_TIM_IMPACT_TMP"
      Tab(2).Control(28)=   "txt_A_TIM_IMPACT_KND(1)"
      Tab(2).Control(29)=   "txt_A_TIM_IMPACT_DSC_CD"
      Tab(2).Control(30)=   "txt_A_TIM_IMPACT_DIR(0)"
      Tab(2).Control(31)=   "txt_A_TIM_IMPACT_DIR(1)"
      Tab(2).Control(32)=   "sdb_TIM_IMPACT_TIM"
      Tab(2).Control(33)=   "ULabel32(5)"
      Tab(2).Control(34)=   "ULabel32(4)"
      Tab(2).Control(35)=   "ULabel32(3)"
      Tab(2).Control(36)=   "ULabel32(2)"
      Tab(2).Control(37)=   "ULabel32(1)"
      Tab(2).Control(38)=   "ULabel32(6)"
      Tab(2).Control(39)=   "ULabel32(11)"
      Tab(2).Control(40)=   "sdb_IMPACT_MIN"
      Tab(2).Control(41)=   "sdb_IMPACT_AVE_MIN"
      Tab(2).Control(42)=   "sdb_IMPACT_RATE_MIN"
      Tab(2).Control(43)=   "sdb_IMPACT_RATE_MAX"
      Tab(2).Control(44)=   "sdb_TIM_IMPACT_MIN"
      Tab(2).Control(45)=   "sdb_TIM_IMPACT_AVE_MIN"
      Tab(2).Control(46)=   "sdb_TIM_IMPACT_RATE_MIN"
      Tab(2).Control(47)=   "sdb_TIM_IMPACT_RATE_MAX"
      Tab(2).Control(48)=   "ULabel32(16)"
      Tab(2).Control(49)=   "ULabel32(17)"
      Tab(2).Control(50)=   "ULabel32(18)"
      Tab(2).Control(51)=   "ULabel32(19)"
      Tab(2).Control(52)=   "ULabel4(32)"
      Tab(2).Control(53)=   "ULabel4(33)"
      Tab(2).Control(54)=   "ULabel4(34)"
      Tab(2).Control(55)=   "ULabel4(35)"
      Tab(2).Control(56)=   "ULabel4(36)"
      Tab(2).Control(57)=   "ULabel4(37)"
      Tab(2).Control(58)=   "ULabel4(38)"
      Tab(2).Control(59)=   "ULabel4(39)"
      Tab(2).Control(60)=   "ULabel4(40)"
      Tab(2).Control(61)=   "ULabel32(8)"
      Tab(2).Control(62)=   "ULabel32(9)"
      Tab(2).Control(63)=   "sdb_IMPACT_MIN_MIN"
      Tab(2).Control(64)=   "sdb_TIM_IMPACT_MIN_MIN"
      Tab(2).Control(65)=   "sdb_A_IMPACT_MIN"
      Tab(2).Control(66)=   "sdb_A_IMPACT_AVE_MIN"
      Tab(2).Control(67)=   "sdb_A_IMPACT_RATE_MIN"
      Tab(2).Control(68)=   "sdb_A_IMPACT_RATE_MAX"
      Tab(2).Control(69)=   "ULabel32(0)"
      Tab(2).Control(70)=   "ULabel32(7)"
      Tab(2).Control(71)=   "ULabel32(10)"
      Tab(2).Control(72)=   "sdb_A_IMPACT_MIN_MIN"
      Tab(2).Control(73)=   "ULabel32(12)"
      Tab(2).Control(74)=   "ULabel32(13)"
      Tab(2).Control(75)=   "ULabel32(14)"
      Tab(2).Control(76)=   "sdb_A_TIM_IMPACT_TIM"
      Tab(2).Control(77)=   "ULabel32(15)"
      Tab(2).Control(78)=   "ULabel32(20)"
      Tab(2).Control(79)=   "ULabel32(21)"
      Tab(2).Control(80)=   "ULabel32(22)"
      Tab(2).Control(81)=   "sdb_A_TIM_IMPACT_MIN"
      Tab(2).Control(82)=   "sdb_A_TIM_IMPACT_AVE_MIN"
      Tab(2).Control(83)=   "sdb_A_TIM_IMPACT_RATE_MIN"
      Tab(2).Control(84)=   "sdb_A_TIM_IMPACT_RATE_MAX"
      Tab(2).Control(85)=   "ULabel32(23)"
      Tab(2).Control(86)=   "ULabel32(24)"
      Tab(2).Control(87)=   "ULabel32(25)"
      Tab(2).Control(88)=   "sdb_A_TIM_IMPACT_MIN_MIN"
      Tab(2).Control(89)=   "ULabel32(26)"
      Tab(2).Control(90)=   "ULabel32(27)"
      Tab(2).Control(91)=   "ULabel32(28)"
      Tab(2).Control(92)=   "ULabel32(29)"
      Tab(2).Control(93)=   "Line49(21)"
      Tab(2).Control(94)=   "Line3(10)"
      Tab(2).Control(95)=   "Line3(11)"
      Tab(2).Control(96)=   "Line3(12)"
      Tab(2).Control(97)=   "Line3(13)"
      Tab(2).Control(98)=   "Line3(14)"
      Tab(2).Control(99)=   "Line49(15)"
      Tab(2).Control(100)=   "Line49(16)"
      Tab(2).ControlCount=   101
      TabCaption(3)   =   "其它"
      TabPicture(3)   =   "AQA0100C.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt_WLD_HARD_UNIT"
      Tab(3).Control(1)=   "txt_DWTT_SMP_CD"
      Tab(3).Control(2)=   "txt_UST_STD_CD(0)"
      Tab(3).Control(3)=   "txt_SSCC_SMP_CD"
      Tab(3).Control(4)=   "txt_JOMINY_SMP_CD"
      Tab(3).Control(5)=   "txt_HIC_SMP_CD"
      Tab(3).Control(6)=   "txt_FOAT_SMP_CD"
      Tab(3).Control(7)=   "txt_RPT_BEND_SMP_CD"
      Tab(3).Control(8)=   "txt_BEND_SMP_CD"
      Tab(3).Control(9)=   "txt_SSCC_SVT_KND(0)"
      Tab(3).Control(10)=   "txt_JOMINY_TYP(0)"
      Tab(3).Control(11)=   "txt_HIC_SVT_KND(0)"
      Tab(3).Control(12)=   "txt_WLD_HARD_TYP(0)"
      Tab(3).Control(13)=   "txt_HARD_TYP(0)"
      Tab(3).Control(14)=   "txt_JOMINY_DSC_CD"
      Tab(3).Control(15)=   "txt_HIC_DSC_CD"
      Tab(3).Control(16)=   "txt_SSCC_DSC_CD"
      Tab(3).Control(17)=   "txt_DWTT_DSC_CD"
      Tab(3).Control(18)=   "txt_FOAT_DSC_CD"
      Tab(3).Control(19)=   "txt_BEND_DSC_CD"
      Tab(3).Control(20)=   "txt_RPT_BEND_DSC_CD"
      Tab(3).Control(21)=   "txt_WLD_HARD_DSC_CD"
      Tab(3).Control(22)=   "txt_WLD_BEND_DSC_CD"
      Tab(3).Control(23)=   "txt_UST_DSC_CD"
      Tab(3).Control(24)=   "txt_HARD_DSC_CD"
      Tab(3).Control(25)=   "txt_DWTT_TMP"
      Tab(3).Control(26)=   "txt_DWTT_TMP_UNIT"
      Tab(3).Control(27)=   "txt_HARD_TYP(1)"
      Tab(3).Control(28)=   "txt_UST_STD_CD(1)"
      Tab(3).Control(29)=   "txt_SSCC_SVT_KND(1)"
      Tab(3).Control(30)=   "txt_JOMINY_TYP(1)"
      Tab(3).Control(31)=   "txt_HIC_SVT_KND(1)"
      Tab(3).Control(32)=   "txt_WLD_HARD_TYP(1)"
      Tab(3).Control(33)=   "txt_UST_GRD"
      Tab(3).Control(34)=   "txt_UST_GRD_NAME"
      Tab(3).Control(35)=   "sdb_HARD_MIN"
      Tab(3).Control(36)=   "sdb_JOMINY_DIST"
      Tab(3).Control(37)=   "ULabel71(9)"
      Tab(3).Control(38)=   "ULabel71(22)"
      Tab(3).Control(39)=   "ULabel71(10)"
      Tab(3).Control(40)=   "ULabel71(11)"
      Tab(3).Control(41)=   "ULabel71(17)"
      Tab(3).Control(42)=   "ULabel71(8)"
      Tab(3).Control(43)=   "ULabel71(12)"
      Tab(3).Control(44)=   "ULabel71(13)"
      Tab(3).Control(45)=   "ULabel71(14)"
      Tab(3).Control(46)=   "ULabel71(15)"
      Tab(3).Control(47)=   "ULabel71(16)"
      Tab(3).Control(48)=   "ULabel71(44)"
      Tab(3).Control(49)=   "ULabel71(41)"
      Tab(3).Control(50)=   "ULabel71(42)"
      Tab(3).Control(51)=   "ULabel71(45)"
      Tab(3).Control(52)=   "ULabel71(40)"
      Tab(3).Control(53)=   "ULabel71(43)"
      Tab(3).Control(54)=   "ULabel99(7)"
      Tab(3).Control(55)=   "sdb_HARD_MAX"
      Tab(3).Control(56)=   "sdb_RPT_BEND_TMS"
      Tab(3).Control(57)=   "sdb_WLD_HARD_MAX"
      Tab(3).Control(58)=   "sdb_DWTT_YP_MIN"
      Tab(3).Control(59)=   "sdb_DWTT_YP_AVE"
      Tab(3).Control(60)=   "sdb_JOMINY_MIN"
      Tab(3).Control(61)=   "sdb_JOMINY_MAX"
      Tab(3).Control(62)=   "sdb_SSCC_YP_MAX"
      Tab(3).Control(63)=   "ULabel71(18)"
      Tab(3).Control(64)=   "ULabel71(19)"
      Tab(3).Control(65)=   "ULabel71(20)"
      Tab(3).Control(66)=   "ULabel71(21)"
      Tab(3).Control(67)=   "ULabel71(23)"
      Tab(3).Control(68)=   "sdb_HIC_CSR_MAX"
      Tab(3).Control(69)=   "sdb_HIC_CLR_MAX"
      Tab(3).Control(70)=   "sdb_HIC_CWR_MAX"
      Tab(3).Control(71)=   "ULabel71(47)"
      Tab(3).Control(72)=   "sdb_SSCC_YP_TIM"
      Tab(3).Control(73)=   "ULabel71(24)"
      Tab(3).Control(74)=   "sdb_BEND_DIA"
      Tab(3).Control(75)=   "sdb_WLD_BEND_DIA"
      Tab(3).Control(76)=   "sdb_BEND_ANGLE"
      Tab(3).Control(77)=   "sdb_WLD_BEND_ANG"
      Tab(3).Control(78)=   "sdb_WLD_HARD_MIN"
      Tab(3).Control(79)=   "UL_HARD_UNIT"
      Tab(3).Control(80)=   "ULabel4(41)"
      Tab(3).Control(81)=   "ULabel4(42)"
      Tab(3).Control(82)=   "ULabel4(43)"
      Tab(3).Control(83)=   "ULabel4(44)"
      Tab(3).Control(84)=   "ULabel4(45)"
      Tab(3).Control(85)=   "ULabel4(46)"
      Tab(3).Control(86)=   "ULabel4(47)"
      Tab(3).Control(87)=   "ULabel4(48)"
      Tab(3).Control(88)=   "ULabel4(49)"
      Tab(3).Control(89)=   "ULabel71(0)"
      Tab(3).Control(90)=   "ULabel71(1)"
      Tab(3).Control(91)=   "Line3(25)"
      Tab(3).Control(92)=   "Line49(14)"
      Tab(3).Control(93)=   "Line49(13)"
      Tab(3).Control(94)=   "Line49(12)"
      Tab(3).Control(95)=   "Line49(11)"
      Tab(3).Control(96)=   "Line49(10)"
      Tab(3).Control(97)=   "Line49(9)"
      Tab(3).Control(98)=   "Line49(8)"
      Tab(3).Control(99)=   "Line49(7)"
      Tab(3).Control(100)=   "Line49(6)"
      Tab(3).Control(101)=   "Line49(5)"
      Tab(3).Control(102)=   "Line3(15)"
      Tab(3).Control(103)=   "Line3(16)"
      Tab(3).Control(104)=   "Line3(17)"
      Tab(3).Control(105)=   "Line3(18)"
      Tab(3).Control(106)=   "Line3(19)"
      Tab(3).ControlCount=   107
      TabCaption(4)   =   "金相检验"
      TabPicture(4)   =   "AQA0100C.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Line49(17)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Line3(24)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Line3(23)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Line3(22)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Line3(21)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Line3(20)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Line49(0)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Line49(27)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Line49(28)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Line49(26)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Line49(4)"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Line49(3)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Line49(2)"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Line49(1)"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "txt_BELT_STR_GRD"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "ULabel87(3)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "ULabel27(0)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "sdb_ACD_DFT_GRD5"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "sdb_ACD_DFT_GRD4"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "ULabel87(0)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "sdb_GRAIN_SIZE_MAX"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "ULabel4(58)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "ULabel4(57)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "ULabel4(56)"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "ULabel4(55)"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "ULabel4(54)"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "ULabel4(53)"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "ULabel4(52)"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "ULabel4(51)"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "ULabel4(50)"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "ULabel27(26)"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).Control(31)=   "sdb_NON_METAL_BGRD4"
      Tab(4).Control(31).Enabled=   0   'False
      Tab(4).Control(32)=   "sdb_NON_METAL_BGRD3"
      Tab(4).Control(32).Enabled=   0   'False
      Tab(4).Control(33)=   "sdb_NON_METAL_BGRD2"
      Tab(4).Control(33).Enabled=   0   'False
      Tab(4).Control(34)=   "sdb_NON_METAL_BGRD1"
      Tab(4).Control(34).Enabled=   0   'False
      Tab(4).Control(35)=   "sdb_NON_METAL_AGRD4"
      Tab(4).Control(35).Enabled=   0   'False
      Tab(4).Control(36)=   "sdb_NON_METAL_AGRD3"
      Tab(4).Control(36).Enabled=   0   'False
      Tab(4).Control(37)=   "sdb_GRAIN_SIZE_TIME"
      Tab(4).Control(37).Enabled=   0   'False
      Tab(4).Control(38)=   "sdb_NON_METAL_AGRD2"
      Tab(4).Control(38).Enabled=   0   'False
      Tab(4).Control(39)=   "sdb_NON_METAL_AGRD1"
      Tab(4).Control(39).Enabled=   0   'False
      Tab(4).Control(40)=   "sdb_ACD_DFT_GRD1"
      Tab(4).Control(40).Enabled=   0   'False
      Tab(4).Control(41)=   "sdb_S_PRINT_DRG"
      Tab(4).Control(41).Enabled=   0   'False
      Tab(4).Control(42)=   "sdb_ACD_DFT_GRD3"
      Tab(4).Control(42).Enabled=   0   'False
      Tab(4).Control(43)=   "sdb_ACD_DFT_GRD2"
      Tab(4).Control(43).Enabled=   0   'False
      Tab(4).Control(44)=   "sdb_GRAIN_SIZE_MIN"
      Tab(4).Control(44).Enabled=   0   'False
      Tab(4).Control(45)=   "sdb_RMV_CAR_MAX"
      Tab(4).Control(45).Enabled=   0   'False
      Tab(4).Control(46)=   "sdb_GRAIN_SIZE_TMP"
      Tab(4).Control(46).Enabled=   0   'False
      Tab(4).Control(47)=   "ULabel87(22)"
      Tab(4).Control(47).Enabled=   0   'False
      Tab(4).Control(48)=   "ULabel87(2)"
      Tab(4).Control(48).Enabled=   0   'False
      Tab(4).Control(49)=   "ULabel87(21)"
      Tab(4).Control(49).Enabled=   0   'False
      Tab(4).Control(50)=   "ULabel87(18)"
      Tab(4).Control(50).Enabled=   0   'False
      Tab(4).Control(51)=   "ULabel87(15)"
      Tab(4).Control(51).Enabled=   0   'False
      Tab(4).Control(52)=   "ULabel87(12)"
      Tab(4).Control(52).Enabled=   0   'False
      Tab(4).Control(53)=   "ULabel87(11)"
      Tab(4).Control(53).Enabled=   0   'False
      Tab(4).Control(54)=   "ULabel87(6)"
      Tab(4).Control(54).Enabled=   0   'False
      Tab(4).Control(55)=   "ULabel87(5)"
      Tab(4).Control(55).Enabled=   0   'False
      Tab(4).Control(56)=   "ULabel87(4)"
      Tab(4).Control(56).Enabled=   0   'False
      Tab(4).Control(57)=   "ULabel87(1)"
      Tab(4).Control(57).Enabled=   0   'False
      Tab(4).Control(58)=   "ULabel27(25)"
      Tab(4).Control(58).Enabled=   0   'False
      Tab(4).Control(59)=   "ULabel27(24)"
      Tab(4).Control(59).Enabled=   0   'False
      Tab(4).Control(60)=   "ULabel27(23)"
      Tab(4).Control(60).Enabled=   0   'False
      Tab(4).Control(61)=   "ULabel27(22)"
      Tab(4).Control(61).Enabled=   0   'False
      Tab(4).Control(62)=   "ULabel27(21)"
      Tab(4).Control(62).Enabled=   0   'False
      Tab(4).Control(63)=   "txt_BELT_STR_DSC_CD"
      Tab(4).Control(63).Enabled=   0   'False
      Tab(4).Control(64)=   "txt_NON_METAL_BCD4(0)"
      Tab(4).Control(64).Enabled=   0   'False
      Tab(4).Control(65)=   "txt_NON_METAL_BCD4(1)"
      Tab(4).Control(65).Enabled=   0   'False
      Tab(4).Control(66)=   "txt_NON_METAL_BCD1(0)"
      Tab(4).Control(66).Enabled=   0   'False
      Tab(4).Control(67)=   "txt_NON_METAL_BCD2(0)"
      Tab(4).Control(67).Enabled=   0   'False
      Tab(4).Control(68)=   "txt_NON_METAL_BCD3(0)"
      Tab(4).Control(68).Enabled=   0   'False
      Tab(4).Control(69)=   "txt_NON_METAL_BCD1(1)"
      Tab(4).Control(69).Enabled=   0   'False
      Tab(4).Control(70)=   "txt_NON_METAL_BCD2(1)"
      Tab(4).Control(70).Enabled=   0   'False
      Tab(4).Control(71)=   "txt_NON_METAL_BCD3(1)"
      Tab(4).Control(71).Enabled=   0   'False
      Tab(4).Control(72)=   "txt_NON_METAL_ACD4(0)"
      Tab(4).Control(72).Enabled=   0   'False
      Tab(4).Control(73)=   "txt_NON_METAL_ACD4(1)"
      Tab(4).Control(73).Enabled=   0   'False
      Tab(4).Control(74)=   "txt_NON_METAL_ACD1(0)"
      Tab(4).Control(74).Enabled=   0   'False
      Tab(4).Control(75)=   "txt_NON_METAL_ACD2(0)"
      Tab(4).Control(75).Enabled=   0   'False
      Tab(4).Control(76)=   "txt_NON_METAL_ACD3(0)"
      Tab(4).Control(76).Enabled=   0   'False
      Tab(4).Control(77)=   "txt_NON_METAL_ACD1(1)"
      Tab(4).Control(77).Enabled=   0   'False
      Tab(4).Control(78)=   "txt_NON_METAL_ACD2(1)"
      Tab(4).Control(78).Enabled=   0   'False
      Tab(4).Control(79)=   "txt_NON_METAL_ACD3(1)"
      Tab(4).Control(79).Enabled=   0   'False
      Tab(4).Control(80)=   "txt_FRACT_GRD4"
      Tab(4).Control(80).Enabled=   0   'False
      Tab(4).Control(81)=   "txt_FRACT_GRD5"
      Tab(4).Control(81).Enabled=   0   'False
      Tab(4).Control(82)=   "txt_FRACT_NAME_CD4(0)"
      Tab(4).Control(82).Enabled=   0   'False
      Tab(4).Control(83)=   "txt_FRACT_NAME_CD5(0)"
      Tab(4).Control(83).Enabled=   0   'False
      Tab(4).Control(84)=   "txt_FRACT_NAME_CD4(1)"
      Tab(4).Control(84).Enabled=   0   'False
      Tab(4).Control(85)=   "txt_FRACT_NAME_CD5(1)"
      Tab(4).Control(85).Enabled=   0   'False
      Tab(4).Control(86)=   "txt_ACD_DFT_TYP4(0)"
      Tab(4).Control(86).Enabled=   0   'False
      Tab(4).Control(87)=   "txt_ACD_DFT_TYP5(0)"
      Tab(4).Control(87).Enabled=   0   'False
      Tab(4).Control(88)=   "txt_ACD_DFT_TYP4(1)"
      Tab(4).Control(88).Enabled=   0   'False
      Tab(4).Control(89)=   "txt_ACD_DFT_TYP5(1)"
      Tab(4).Control(89).Enabled=   0   'False
      Tab(4).Control(90)=   "txt_GRAIN_SIZE_TMP_UNIT"
      Tab(4).Control(90).Enabled=   0   'False
      Tab(4).Control(91)=   "txt_FRACT_GRD3"
      Tab(4).Control(91).Enabled=   0   'False
      Tab(4).Control(92)=   "txt_FRACT_GRD2"
      Tab(4).Control(92).Enabled=   0   'False
      Tab(4).Control(93)=   "txt_FRACT_GRD1"
      Tab(4).Control(93).Enabled=   0   'False
      Tab(4).Control(94)=   "txt_RMV_CAR_DSC_CD"
      Tab(4).Control(94).Enabled=   0   'False
      Tab(4).Control(95)=   "txt_NON_METAL_DSC_CD"
      Tab(4).Control(95).Enabled=   0   'False
      Tab(4).Control(96)=   "txt_FRACT_DSC_CD"
      Tab(4).Control(96).Enabled=   0   'False
      Tab(4).Control(97)=   "txt_ACD_DSC_CD"
      Tab(4).Control(97).Enabled=   0   'False
      Tab(4).Control(98)=   "txt_S_PRINT_DSC_CD"
      Tab(4).Control(98).Enabled=   0   'False
      Tab(4).Control(99)=   "txt_GRAIN_SIZE_DSC_CD"
      Tab(4).Control(99).Enabled=   0   'False
      Tab(4).Control(100)=   "txt_FRACT_NAME_CD3(1)"
      Tab(4).Control(100).Enabled=   0   'False
      Tab(4).Control(101)=   "txt_FRACT_NAME_CD2(1)"
      Tab(4).Control(101).Enabled=   0   'False
      Tab(4).Control(102)=   "txt_FRACT_NAME_CD1(1)"
      Tab(4).Control(102).Enabled=   0   'False
      Tab(4).Control(103)=   "txt_FRACT_NAME_CD3(0)"
      Tab(4).Control(103).Enabled=   0   'False
      Tab(4).Control(104)=   "txt_FRACT_NAME_CD2(0)"
      Tab(4).Control(104).Enabled=   0   'False
      Tab(4).Control(105)=   "txt_FRACT_NAME_CD1(0)"
      Tab(4).Control(105).Enabled=   0   'False
      Tab(4).Control(106)=   "txt_ACD_DFT_TYP3(1)"
      Tab(4).Control(106).Enabled=   0   'False
      Tab(4).Control(107)=   "txt_ACD_DFT_TYP2(1)"
      Tab(4).Control(107).Enabled=   0   'False
      Tab(4).Control(108)=   "txt_ACD_DFT_TYP1(1)"
      Tab(4).Control(108).Enabled=   0   'False
      Tab(4).Control(109)=   "txt_ACD_DFT_TYP3(0)"
      Tab(4).Control(109).Enabled=   0   'False
      Tab(4).Control(110)=   "txt_ACD_DFT_TYP2(0)"
      Tab(4).Control(110).Enabled=   0   'False
      Tab(4).Control(111)=   "txt_ACD_DFT_TYP1(0)"
      Tab(4).Control(111).Enabled=   0   'False
      Tab(4).Control(112)=   "txt_RMV_CAR_SMP_CD"
      Tab(4).Control(112).Enabled=   0   'False
      Tab(4).Control(113)=   "txt_RMV_CAR_TYP(1)"
      Tab(4).Control(113).Enabled=   0   'False
      Tab(4).Control(114)=   "txt_RMV_CAR_TYP(0)"
      Tab(4).Control(114).Enabled=   0   'False
      Tab(4).Control(115)=   "txt_GRAIN_SIZE_MTH(1)"
      Tab(4).Control(115).Enabled=   0   'False
      Tab(4).Control(116)=   "txt_GRAIN_SIZE_MTH(0)"
      Tab(4).Control(116).Enabled=   0   'False
      Tab(4).Control(117)=   "txt_FRACT_SMP_CD"
      Tab(4).Control(117).Enabled=   0   'False
      Tab(4).Control(118)=   "txt_NON_METAL_SMP_CD"
      Tab(4).Control(118).Enabled=   0   'False
      Tab(4).Control(119)=   "txt_NON_METAL_TYP(1)"
      Tab(4).Control(119).Enabled=   0   'False
      Tab(4).Control(120)=   "txt_NON_METAL_TYP(0)"
      Tab(4).Control(120).Enabled=   0   'False
      Tab(4).ControlCount=   121
      Begin VB.TextBox txt_RA_DIR_NAME 
         Height          =   300
         Left            =   -70380
         TabIndex        =   256
         Top             =   2670
         Width           =   1470
      End
      Begin VB.TextBox txt_RA_DIR_CD 
         Height          =   300
         Left            =   -70890
         TabIndex        =   255
         Top             =   2670
         Width           =   495
      End
      Begin VB.TextBox txt_YR_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   252
         Top             =   4546
         Width           =   900
      End
      Begin VB.TextBox txt_TIM_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   164
         Top             =   4740
         Width           =   1455
      End
      Begin VB.TextBox txt_TIM_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   163
         Top             =   4740
         Width           =   495
      End
      Begin VB.TextBox txt_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   162
         Top             =   1255
         Width           =   1470
      End
      Begin VB.TextBox txt_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "宋体"
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
         Top             =   1255
         Width           =   495
      End
      Begin VB.TextBox txt_NON_METAL_TYP 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   4110
         MaxLength       =   1
         TabIndex        =   160
         Top             =   4950
         Width           =   495
      End
      Begin VB.TextBox txt_NON_METAL_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   4620
         MaxLength       =   80
         TabIndex        =   159
         Top             =   4950
         Width           =   1470
      End
      Begin VB.TextBox txt_NON_METAL_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2085
         MaxLength       =   9
         TabIndex        =   158
         Top             =   4950
         Width           =   1980
      End
      Begin VB.TextBox txt_FRACT_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2085
         MaxLength       =   9
         TabIndex        =   157
         Top             =   3480
         Width           =   1980
      End
      Begin VB.TextBox txt_GRAIN_SIZE_MTH 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   4110
         MaxLength       =   1
         TabIndex        =   156
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_GRAIN_SIZE_MTH 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   4620
         MaxLength       =   80
         TabIndex        =   155
         Top             =   1200
         Width           =   1470
      End
      Begin VB.TextBox txt_RMV_CAR_TYP 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   4110
         MaxLength       =   1
         TabIndex        =   154
         Top             =   810
         Width           =   495
      End
      Begin VB.TextBox txt_RMV_CAR_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   4620
         MaxLength       =   80
         TabIndex        =   153
         Top             =   810
         Width           =   1470
      End
      Begin VB.TextBox txt_RMV_CAR_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2085
         MaxLength       =   9
         TabIndex        =   152
         Top             =   810
         Width           =   1980
      End
      Begin VB.TextBox txt_WLD_HARD_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -61710
         MaxLength       =   4
         TabIndex        =   151
         Top             =   2370
         Width           =   900
      End
      Begin VB.TextBox txt_DWTT_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   150
         Top             =   7230
         Width           =   1980
      End
      Begin VB.TextBox txt_UST_STD_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   149
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox txt_SSCC_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   148
         Top             =   6630
         Width           =   1980
      End
      Begin VB.TextBox txt_JOMINY_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   147
         Top             =   4635
         Width           =   1980
      End
      Begin VB.TextBox txt_HIC_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   146
         Top             =   5220
         Width           =   1980
      End
      Begin VB.TextBox txt_FOAT_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   145
         Top             =   4080
         Width           =   1980
      End
      Begin VB.TextBox txt_RPT_BEND_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   144
         Top             =   1815
         Width           =   1980
      End
      Begin VB.TextBox txt_BEND_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   143
         Top             =   1305
         Width           =   1980
      End
      Begin VB.TextBox txt_SSCC_SVT_KND 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   142
         Top             =   6630
         Width           =   495
      End
      Begin VB.TextBox txt_JOMINY_TYP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   141
         Top             =   4635
         Width           =   495
      End
      Begin VB.TextBox txt_HIC_SVT_KND 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   140
         Top             =   5220
         Width           =   495
      End
      Begin VB.TextBox txt_WLD_HARD_TYP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   139
         Top             =   2370
         Width           =   495
      End
      Begin VB.TextBox txt_HARD_TYP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   138
         Top             =   810
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   137
         Tag             =   "Q0002"
         Top             =   4080
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   136
         Tag             =   "Q0002"
         Top             =   5170
         Width           =   900
      End
      Begin VB.TextBox txt_ACD_DFT_TYP1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   135
         Top             =   1980
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   134
         Top             =   2265
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   133
         Top             =   2550
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   132
         Top             =   1980
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   131
         Top             =   2265
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   130
         Top             =   2550
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   129
         Top             =   3480
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   128
         Top             =   3765
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   127
         Top             =   4050
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   126
         Top             =   3480
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   125
         Top             =   3765
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   124
         Top             =   4050
         Width           =   1605
      End
      Begin VB.TextBox txt_GRAIN_SIZE_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14220
         MaxLength       =   1
         TabIndex        =   123
         Top             =   1215
         Width           =   900
      End
      Begin VB.TextBox txt_S_PRINT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14220
         MaxLength       =   1
         TabIndex        =   122
         Top             =   1590
         Width           =   900
      End
      Begin VB.TextBox txt_ACD_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   14220
         MaxLength       =   1
         TabIndex        =   121
         Top             =   1980
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14220
         MaxLength       =   1
         TabIndex        =   120
         Top             =   3480
         Width           =   900
      End
      Begin VB.TextBox txt_NON_METAL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14220
         MaxLength       =   1
         TabIndex        =   119
         Top             =   4950
         Width           =   900
      End
      Begin VB.TextBox txt_RMV_CAR_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14220
         MaxLength       =   1
         TabIndex        =   118
         Top             =   810
         Width           =   900
      End
      Begin VB.TextBox txt_JOMINY_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   117
         Top             =   4635
         Width           =   900
      End
      Begin VB.TextBox txt_HIC_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   116
         Top             =   5220
         Width           =   900
      End
      Begin VB.TextBox txt_SSCC_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   115
         Top             =   6600
         Width           =   900
      End
      Begin VB.TextBox txt_DWTT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   114
         Top             =   7230
         Width           =   900
      End
      Begin VB.TextBox txt_FOAT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   113
         Top             =   4080
         Width           =   900
      End
      Begin VB.TextBox txt_BEND_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   112
         Top             =   1305
         Width           =   900
      End
      Begin VB.TextBox txt_RPT_BEND_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   111
         Top             =   1815
         Width           =   900
      End
      Begin VB.TextBox txt_WLD_HARD_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   110
         Top             =   2370
         Width           =   900
      End
      Begin VB.TextBox txt_WLD_BEND_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   109
         Top             =   2955
         Width           =   900
      End
      Begin VB.TextBox txt_UST_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   108
         Top             =   3540
         Width           =   900
      End
      Begin VB.TextBox txt_HARD_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   107
         Top             =   810
         Width           =   900
      End
      Begin VB.TextBox txt_TIM_IMPACT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   106
         Top             =   4305
         Width           =   900
      End
      Begin VB.TextBox txt_IMPACT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   105
         Top             =   810
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_TS_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   104
         Tag             =   "Q0002"
         Top             =   1900
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   103
         Tag             =   "Q0002"
         Top             =   2990
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_SP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   102
         Tag             =   "Q0002"
         Top             =   6260
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_YP_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   101
         Tag             =   "Q0002"
         Top             =   810
         Width           =   900
      End
      Begin VB.TextBox txt_SP_EL_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   100
         Top             =   7350
         Width           =   495
      End
      Begin VB.TextBox txt_EL_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         Top             =   3612
         Width           =   495
      End
      Begin VB.TextBox txt_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         Tag             =   "N"
         Top             =   3612
         Width           =   1470
      End
      Begin VB.TextBox txt_SP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   97
         Tag             =   "N"
         Top             =   7350
         Width           =   1470
      End
      Begin VB.TextBox txt_SG_EL_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   96
         Top             =   6414
         Width           =   495
      End
      Begin VB.TextBox txt_SG_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   95
         Tag             =   "N"
         Top             =   6414
         Width           =   1470
      End
      Begin VB.TextBox txt_SNPP_EL_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   94
         Top             =   5480
         Width           =   495
      End
      Begin VB.TextBox txt_SNPP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   93
         Tag             =   "N"
         Top             =   5480
         Width           =   1470
      End
      Begin VB.TextBox txt_SP_EL_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   92
         Top             =   7350
         Width           =   1980
      End
      Begin VB.TextBox txt_TENCIL_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   91
         Top             =   810
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_SP_EL_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   90
         Top             =   6260
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_SP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   89
         Top             =   6260
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   88
         Top             =   5170
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   87
         Top             =   5170
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_EL_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   86
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   85
         Top             =   4080
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_SP_EL_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   84
         Top             =   6260
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_TENCIL_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   83
         Top             =   810
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_TENCIL_TMP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   82
         Top             =   810
         Width           =   1485
      End
      Begin VB.TextBox txt_HGT_TENCIL_TMP_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   81
         Top             =   810
         Width           =   510
      End
      Begin VB.TextBox txt_IMPACT_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   80
         Top             =   810
         Width           =   1980
      End
      Begin VB.TextBox txt_IMPACT_KND 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   79
         Top             =   810
         Width           =   495
      End
      Begin VB.TextBox txt_FRACT_GRD1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12360
         MaxLength       =   1
         TabIndex        =   78
         Top             =   3480
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_GRD2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12360
         MaxLength       =   1
         TabIndex        =   77
         Top             =   3765
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_GRD3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12360
         MaxLength       =   1
         TabIndex        =   76
         Top             =   4050
         Width           =   900
      End
      Begin VB.TextBox txt_GRAIN_SIZE_TMP_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7635
         MaxLength       =   1
         TabIndex        =   75
         Top             =   1200
         Width           =   510
      End
      Begin VB.TextBox txt_DWTT_TMP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   74
         Top             =   7230
         Width           =   1485
      End
      Begin VB.TextBox txt_DWTT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67350
         MaxLength       =   1
         TabIndex        =   73
         Top             =   7230
         Width           =   510
      End
      Begin VB.TextBox txt_HARD_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   72
         Top             =   810
         Width           =   1470
      End
      Begin VB.TextBox txt_UST_STD_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   71
         Top             =   3540
         Width           =   3525
      End
      Begin VB.TextBox txt_SSCC_SVT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   70
         Top             =   6630
         Width           =   1470
      End
      Begin VB.TextBox txt_JOMINY_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   69
         Top             =   4635
         Width           =   1470
      End
      Begin VB.TextBox txt_HIC_SVT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   68
         Top             =   5220
         Width           =   1470
      End
      Begin VB.TextBox txt_WLD_HARD_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   67
         Top             =   2370
         Width           =   1470
      End
      Begin VB.TextBox txt_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   66
         Top             =   810
         Width           =   1470
      End
      Begin VB.TextBox txt_TIM_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   65
         Top             =   4305
         Width           =   1455
      End
      Begin VB.TextBox txt_TIM_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   64
         Top             =   4305
         Width           =   1485
      End
      Begin VB.TextBox txt_TIM_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   63
         Top             =   4305
         Width           =   510
      End
      Begin VB.TextBox txt_TIM_IMPACT_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   62
         Top             =   4305
         Width           =   1980
      End
      Begin VB.TextBox txt_TIM_IMPACT_KND 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   61
         Top             =   4305
         Width           =   495
      End
      Begin VB.TextBox txt_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   60
         Top             =   810
         Width           =   1485
      End
      Begin VB.TextBox txt_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   59
         Top             =   810
         Width           =   510
      End
      Begin VB.TextBox txt_UST_GRD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   58
         Tag             =   "txt_UST_GRD_NAME,Q0053"
         Top             =   3540
         Width           =   900
      End
      Begin VB.TextBox txt_YP_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   57
         Top             =   810
         Width           =   900
      End
      Begin VB.TextBox txt_SP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   56
         Top             =   7350
         Width           =   900
      End
      Begin VB.TextBox txt_SG_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   55
         Top             =   6414
         Width           =   900
      End
      Begin VB.TextBox txt_SNPP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   54
         Top             =   5480
         Width           =   900
      End
      Begin VB.TextBox txt_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   53
         Top             =   3612
         Width           =   900
      End
      Begin VB.TextBox txt_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   52
         Top             =   2678
         Width           =   900
      End
      Begin VB.TextBox txt_TS_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   51
         Top             =   1744
         Width           =   900
      End
      Begin VB.TextBox txt_UST_GRD_NAME 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   50
         Top             =   3900
         Width           =   900
      End
      Begin VB.TextBox txt_A_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   49
         Top             =   2550
         Width           =   510
      End
      Begin VB.TextBox txt_A_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   48
         Top             =   2550
         Width           =   1485
      End
      Begin VB.TextBox txt_A_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   47
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_A_IMPACT_KND 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   46
         Top             =   2550
         Width           =   495
      End
      Begin VB.TextBox txt_A_IMPACT_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72930
         MaxLength       =   9
         TabIndex        =   45
         Top             =   2550
         Width           =   1980
      End
      Begin VB.TextBox txt_A_IMPACT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60795
         MaxLength       =   1
         TabIndex        =   44
         Top             =   2550
         Width           =   900
      End
      Begin VB.TextBox txt_A_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   43
         Top             =   2995
         Width           =   495
      End
      Begin VB.TextBox txt_A_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   42
         Top             =   2995
         Width           =   1470
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_KND 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   41
         Top             =   6030
         Width           =   495
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_SMP_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   40
         Top             =   6030
         Width           =   1980
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   39
         Top             =   6030
         Width           =   510
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   38
         Top             =   6030
         Width           =   1485
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   37
         Top             =   6030
         Width           =   1455
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   36
         Top             =   6030
         Width           =   900
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   35
         Top             =   6465
         Width           =   495
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   34
         Top             =   6465
         Width           =   1455
      End
      Begin VB.TextBox txt_ACD_DFT_TYP5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   33
         Top             =   3135
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   32
         Top             =   2850
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   31
         Top             =   3135
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   30
         Top             =   2850
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   29
         Top             =   4605
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   28
         Top             =   4320
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   27
         Top             =   4605
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   2
         TabIndex        =   26
         Top             =   4320
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_GRD5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12360
         MaxLength       =   1
         TabIndex        =   25
         Top             =   4605
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_GRD4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12360
         MaxLength       =   1
         TabIndex        =   24
         Top             =   4320
         Width           =   900
      End
      Begin VB.TextBox txt_NON_METAL_ACD3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   23
         Top             =   5520
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   22
         Top             =   5235
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   21
         Top             =   4950
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   20
         Top             =   5520
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   19
         Top             =   5235
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   18
         Top             =   4950
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   17
         Top             =   5790
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   16
         Top             =   5790
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   15
         Top             =   6720
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   14
         Top             =   6435
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   13
         Top             =   6150
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   12
         Top             =   6720
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   11
         Top             =   6435
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   10
         Top             =   6150
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9750
         MaxLength       =   80
         TabIndex        =   9
         Top             =   6990
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   8
         Top             =   6990
         Width           =   435
      End
      Begin VB.TextBox txt_BELT_STR_DSC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14220
         MaxLength       =   1
         TabIndex        =   7
         Top             =   7350
         Width           =   900
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   1
         Left            =   -72915
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "取样代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HARD_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   165
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_DIST 
         Height          =   300
         Left            =   -65910
         TabIndex        =   166
         Top             =   4635
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_TIM 
         Height          =   300
         Left            =   -65520
         TabIndex        =   167
         Top             =   4305
         Width           =   1875
         _Version        =   262145
         _ExtentX        =   3307
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   0
         Left            =   -74970
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   21
         Left            =   -74970
         Top             =   1744
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "抗拉强度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   20
         Left            =   -74970
         Top             =   2678
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断面收缩率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   19
         Left            =   -74970
         Top             =   3612
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断后伸长率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   18
         Left            =   -74970
         Top             =   5480
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "规定非比例伸长应力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   17
         Left            =   -74970
         Top             =   6414
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "规定总伸长应力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   16
         Left            =   -74970
         Top             =   7350
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "规定残余伸长应力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   2
         Left            =   -70890
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "其它代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   3
         Left            =   -68865
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验温度"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   4
         Left            =   -66810
         Top             =   450
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "其它项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   5
         Left            =   -63570
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "下限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   7
         Left            =   -62640
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "上限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   6
         Left            =   -61710
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "单位"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   8
         Left            =   -60780
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "判定"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   22
         Left            =   -74970
         Top             =   810
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "屈服强度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   3
         Left            =   -74970
         Top             =   1900
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "抗拉强度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   4
         Left            =   -74970
         Top             =   2990
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断面收缩率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   5
         Left            =   -74970
         Top             =   4080
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断后伸长率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   6
         Left            =   -74970
         Top             =   5170
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "规定非比例伸长应力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   7
         Left            =   -74970
         Top             =   6260
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "规定残余伸长应力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   2
         Left            =   -74970
         Top             =   810
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "屈服强度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   168
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   169
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -61710
         Top             =   810
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   170
         Top             =   1900
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   171
         Top             =   1900
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -61710
         Top             =   1900
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   172
         Top             =   2990
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   173
         Top             =   2990
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -61710
         Top             =   2990
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   174
         Top             =   6260
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   175
         Top             =   6260
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -61710
         Top             =   6260
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
         Height          =   300
         Index           =   5
         Left            =   -74970
         Top             =   1700
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "平均"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   4
         Left            =   -74970
         Top             =   2145
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断面纤维率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   3
         Left            =   -74970
         Top             =   4305
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "时效冲击试验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   2
         Left            =   -74970
         Top             =   5175
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "平均"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   1
         Left            =   -74970
         Top             =   5610
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断面纤维率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   6
         Left            =   -74970
         Top             =   810
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "冲击试验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   9
         Left            =   -74970
         Top             =   1305
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "弯曲试验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   22
         Left            =   -74970
         Top             =   1815
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "反复弯曲"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   10
         Left            =   -74970
         Top             =   2370
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "焊缝硬度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   11
         Left            =   -74970
         Top             =   2955
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "焊缝弯曲"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   17
         Left            =   -74970
         Top             =   7230
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "重力撕裂试验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   8
         Left            =   -74970
         Top             =   810
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "硬度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   11
         Left            =   -66810
         Top             =   4305
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "时效时间"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   176
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   177
         Top             =   1700
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   178
         Top             =   2145
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   179
         Top             =   2145
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   180
         Top             =   4305
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   181
         Top             =   5175
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   182
         Top             =   5610
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   183
         Top             =   5610
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   810
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J"
         Alignment       =   1
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
         Height          =   300
         Index           =   17
         Left            =   -61710
         Top             =   2145
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
         Height          =   300
         Index           =   18
         Left            =   -61710
         Top             =   4305
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J/cm2"
         Alignment       =   1
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
         Height          =   300
         Index           =   19
         Left            =   -61710
         Top             =   5610
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   12
         Left            =   -74970
         Top             =   3540
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "超声波探伤（UST）"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   13
         Left            =   -74970
         Top             =   4080
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "锻平"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   14
         Left            =   -74970
         Top             =   4635
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "淬透性"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   15
         Left            =   -74970
         Top             =   5220
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "抗氢裂能力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   16
         Left            =   -74970
         Top             =   6630
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "硫化物腐蚀裂纹"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   44
         Left            =   -66780
         Top             =   4635
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "距端面"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   41
         Left            =   -66780
         Top             =   1305
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "弯心直径"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   42
         Left            =   -66780
         Top             =   2955
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "弯心直径"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   45
         Left            =   -66780
         Top             =   5220
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "CSR"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   40
         Left            =   -65160
         Top             =   1305
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "弯曲角度"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   43
         Left            =   -65160
         Top             =   2955
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "弯曲角度"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel99 
         Height          =   300
         Index           =   7
         Left            =   -64680
         Top             =   4635
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   529
         Caption         =   "mm"
         Alignment       =   0
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
      Begin CSTextLibCtl.sidbEdit sdb_HARD_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   184
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -63570
         TabIndex        =   185
         Top             =   1815
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   186
         Top             =   2370
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_MIN 
         Height          =   300
         Left            =   -65865
         TabIndex        =   187
         Top             =   7230
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   188
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
            Name            =   "宋体"
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
         TabIndex        =   189
         Top             =   4635
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   190
         Top             =   4635
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   191
         Top             =   6600
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   3540
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "级"
         Alignment       =   1
         BackgroundStyle =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   19
         Left            =   -61710
         Top             =   7230
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   20
         Left            =   -61710
         Top             =   1815
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "次"
         Alignment       =   1
         BackgroundStyle =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   21
         Left            =   -61710
         Top             =   6600
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   21
         Left            =   30
         Top             =   1200
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "晶粒度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   22
         Left            =   30
         Top             =   1590
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "硫印"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   23
         Left            =   30
         Top             =   1980
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "酸浸检验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   24
         Left            =   30
         Top             =   3480
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断口检验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   25
         Left            =   30
         Top             =   4950
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "非金属夹杂"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   1
         Left            =   8190
         Top             =   4950
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "粗系"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   270
         Index           =   4
         Left            =   13290
         Top             =   1980
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   476
         Caption         =   "级"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   5
         Left            =   13290
         Top             =   1590
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "级"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   270
         Index           =   6
         Left            =   8190
         Top             =   1980
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   476
         Caption         =   "缺陷名称"
         Alignment       =   1
         BackColor       =   12632256
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.74
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
         Left            =   8190
         Top             =   3480
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "断口名称"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   12
         Left            =   13290
         Top             =   3480
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "级"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   15
         Left            =   13290
         Top             =   1215
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "级"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   18
         Left            =   13290
         Top             =   810
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "mm"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   21
         Left            =   8220
         Top             =   1200
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "保温时间"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   2
         Left            =   8190
         Top             =   6150
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Caption         =   "细系"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   22
         Left            =   13290
         Top             =   4950
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "级"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   23
         Left            =   -66780
         Top             =   5610
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "CLR"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   192
         Top             =   5220
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   193
         Top             =   5610
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   194
         Top             =   6000
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -66780
         Top             =   6600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "时间"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_SSCC_YP_TIM 
         Height          =   300
         Left            =   -65910
         TabIndex        =   195
         Top             =   6600
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -66780
         Top             =   6000
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "CWR"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_BEND_DIA 
         Height          =   300
         Left            =   -65910
         TabIndex        =   196
         Top             =   1305
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -65910
         TabIndex        =   197
         Top             =   2955
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -64290
         TabIndex        =   198
         Top             =   1305
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -64290
         TabIndex        =   199
         Top             =   2955
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   200
         Top             =   5170
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   201
         Top             =   5170
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -61710
         Top             =   5170
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   202
         Top             =   4080
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   203
         Top             =   4080
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -61710
         Top             =   4080
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_WLD_HARD_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   204
         Top             =   2370
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -61710
         Top             =   810
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
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_TMP 
         Height          =   300
         Left            =   6135
         TabIndex        =   205
         Top             =   1200
         Width           =   1485
         _Version        =   262145
         _ExtentX        =   2619
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin CSTextLibCtl.sidbEdit sdb_RMV_CAR_MAX 
         Height          =   300
         Left            =   12360
         TabIndex        =   206
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   11430
         TabIndex        =   207
         Top             =   1200
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_DFT_GRD2 
         Height          =   270
         Left            =   12360
         TabIndex        =   208
         Top             =   2265
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   209
         Top             =   2550
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   210
         Top             =   1590
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   211
         Top             =   1980
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_AGRD1 
         Height          =   270
         Left            =   12360
         TabIndex        =   212
         Top             =   4950
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   213
         Top             =   5235
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_TIME 
         Height          =   300
         Left            =   9300
         TabIndex        =   214
         Top             =   1200
         Width           =   1425
         _Version        =   262145
         _ExtentX        =   2514
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_AGRD3 
         Height          =   270
         Left            =   12360
         TabIndex        =   215
         Top             =   5520
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   216
         Top             =   5790
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   217
         Top             =   6150
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   218
         Top             =   6435
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   219
         Top             =   6720
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   220
         Top             =   6990
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin CSTextLibCtl.sidbEdit sdb_YP_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   221
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   222
         Top             =   810
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   810
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_TS_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   223
         Top             =   1744
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   224
         Top             =   1744
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   1744
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_RA_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   225
         Top             =   2678
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   226
         Top             =   2678
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   2678
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_EL_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   227
         Top             =   3612
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   228
         Top             =   3612
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   3612
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_SNPP_EL_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   229
         Top             =   5480
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   230
         Top             =   5480
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   5480
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_SG_EL_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   231
         Top             =   6414
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   232
         Top             =   6414
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   6414
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_SP_EL_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   233
         Top             =   7350
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   234
         Top             =   7350
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   7350
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   26
         Left            =   30
         Top             =   840
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "脱碳层"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   23
         Left            =   -72915
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "取样代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   24
         Left            =   -74970
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   25
         Left            =   -70890
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "其它代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   26
         Left            =   -68865
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验温度"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   27
         Left            =   -66810
         Top             =   450
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "其它项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   28
         Left            =   -63570
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "下限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   29
         Left            =   -62640
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "上限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   30
         Left            =   -61710
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "单位"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   31
         Left            =   -60780
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "判定"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   32
         Left            =   -72915
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "取样代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   33
         Left            =   -74970
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   34
         Left            =   -70890
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "其它代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   35
         Left            =   -68865
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验温度"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   36
         Left            =   -66810
         Top             =   450
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "其它项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   37
         Left            =   -63570
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "下限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   38
         Left            =   -62640
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "上限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   39
         Left            =   -61710
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "单位"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   40
         Left            =   -60780
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "判定"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   41
         Left            =   -72915
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "取样代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   42
         Left            =   -74970
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   43
         Left            =   -70890
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "其它代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   44
         Left            =   -68865
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验温度"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   45
         Left            =   -66810
         Top             =   450
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "其它项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   46
         Left            =   -63570
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "下限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   47
         Left            =   -62640
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "上限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   48
         Left            =   -61710
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "单位"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   49
         Left            =   -60780
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "判定"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   50
         Left            =   2085
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "取样代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   51
         Left            =   30
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   52
         Left            =   4110
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "其它代码"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   53
         Left            =   6135
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "试验温度"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   54
         Left            =   8190
         Top             =   450
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "其它项目"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   55
         Left            =   11430
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "下限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   56
         Left            =   12360
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "上限"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   57
         Left            =   13290
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "单位"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   58
         Left            =   14220
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "判定"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   8
         Left            =   -66810
         Top             =   1255
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "最小下限"
         Alignment       =   1
         BackColor       =   12632256
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
         Height          =   300
         Index           =   9
         Left            =   -66810
         Top             =   4740
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "最小下限"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_MIN_MIN 
         Height          =   300
         Left            =   -65520
         TabIndex        =   235
         Top             =   1255
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   236
         Top             =   4740
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   237
         Top             =   2550
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -63585
         TabIndex        =   238
         Top             =   3440
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   239
         Top             =   3885
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   240
         Top             =   3885
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   2550
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J"
         Alignment       =   1
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
         Height          =   300
         Index           =   7
         Left            =   -61725
         Top             =   3885
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
         Height          =   300
         Index           =   10
         Left            =   -66810
         Top             =   2995
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "最小下限"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_MIN_MIN 
         Height          =   300
         Left            =   -65520
         TabIndex        =   241
         Top             =   2995
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   3440
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "平均"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   13
         Left            =   -74970
         Top             =   3885
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断面纤维率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   14
         Left            =   -74970
         Top             =   2550
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "追加冲击试验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_MAX 
         Height          =   300
         Left            =   12360
         TabIndex        =   242
         Top             =   1200
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_TIM 
         Height          =   300
         Left            =   -65520
         TabIndex        =   243
         Top             =   6030
         Width           =   1875
         _Version        =   262145
         _ExtentX        =   3307
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   6030
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "追加时效冲击试验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   20
         Left            =   -74970
         Top             =   6900
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "平均"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   21
         Left            =   -74970
         Top             =   7335
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "断面纤维率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   22
         Left            =   -66810
         Top             =   6030
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "时效时间"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   244
         Top             =   6030
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   245
         Top             =   6915
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   246
         Top             =   7335
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   247
         Top             =   7335
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   6030
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J/cm2"
         Alignment       =   1
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
         Height          =   300
         Index           =   24
         Left            =   -61710
         Top             =   7335
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
         Height          =   300
         Index           =   25
         Left            =   -66810
         Top             =   6465
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "最小下限"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_MIN_MIN 
         Height          =   300
         Left            =   -65520
         TabIndex        =   248
         Top             =   6465
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   0
         Left            =   4110
         Top             =   3480
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   $"AQA0100C.frx":008C
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_DFT_GRD4 
         Height          =   270
         Left            =   12360
         TabIndex        =   249
         Top             =   2850
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
         Left            =   12360
         TabIndex        =   250
         Top             =   3135
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   ""
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   0
         Left            =   30
         Top             =   7350
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "带状组织"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   3
         Left            =   13290
         Top             =   7350
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "级"
         Alignment       =   1
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
         Height          =   300
         Index           =   26
         Left            =   -74970
         Top             =   1255
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   27
         Left            =   -74970
         Top             =   2995
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   28
         Left            =   -74970
         Top             =   4740
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   29
         Left            =   -74970
         Top             =   6465
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Height          =   300
         Index           =   59
         Left            =   -74985
         Top             =   4546
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "屈强比"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_YR_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   253
         Top             =   4546
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   -62640
         TabIndex        =   254
         Top             =   4546
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Top             =   4546
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "％"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   0
         Left            =   -66795
         Top             =   7230
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "最小下限"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   1
         Left            =   -74955
         Top             =   7605
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "平均"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit txt_BELT_STR_GRD 
         Height          =   300
         Left            =   12360
         TabIndex        =   257
         Top             =   7350
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
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
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   25
         X1              =   -68880
         X2              =   -68880
         Y1              =   4005
         Y2              =   8100
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   21
         X1              =   -75000
         X2              =   -59820
         Y1              =   4230
         Y2              =   4230
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   14
         X1              =   -75000
         X2              =   -59850
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   -75000
         X2              =   -59850
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   -75000
         X2              =   -59850
         Y1              =   1695
         Y2              =   1695
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   -75000
         X2              =   -59850
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   -75000
         X2              =   -59850
         Y1              =   3390
         Y2              =   3390
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   -75000
         X2              =   -59850
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   -75000
         X2              =   -59850
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   -75000
         X2              =   -59850
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   -75000
         X2              =   -59850
         Y1              =   7065
         Y2              =   7065
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   -75000
         X2              =   -59850
         Y1              =   6450
         Y2              =   6450
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   0
         X2              =   15150
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   0
         X2              =   15150
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   0
         X2              =   15150
         Y1              =   3435
         Y2              =   3435
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   0
         X2              =   15150
         Y1              =   4905
         Y2              =   4905
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   26
         X1              =   8160
         X2              =   15150
         Y1              =   6105
         Y2              =   6105
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   28
         X1              =   9255
         X2              =   9255
         Y1              =   1935
         Y2              =   7305
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   27
         X1              =   0
         X2              =   15150
         Y1              =   7305
         Y2              =   7305
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   15150
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   -72945
         X2              =   -72945
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   -70920
         X2              =   -70920
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   -68895
         X2              =   -68895
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   -66840
         X2              =   -66840
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   -63600
         X2              =   -63600
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   -63600
         X2              =   -63600
         Y1              =   450
         Y2              =   8100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   -66840
         X2              =   -66840
         Y1              =   450
         Y2              =   8100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   -68895
         X2              =   -68895
         Y1              =   450
         Y2              =   8100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   -70920
         X2              =   -70920
         Y1              =   450
         Y2              =   8100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   14
         X1              =   -72945
         X2              =   -72945
         Y1              =   450
         Y2              =   8100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   -63600
         X2              =   -63600
         Y1              =   450
         Y2              =   8130
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   -66840
         X2              =   -66840
         Y1              =   450
         Y2              =   8100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   -68895
         X2              =   -68895
         Y1              =   450
         Y2              =   3390
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   18
         X1              =   -70920
         X2              =   -70920
         Y1              =   450
         Y2              =   8115
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   19
         X1              =   -72945
         X2              =   -72945
         Y1              =   450
         Y2              =   8115
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   20
         X1              =   11400
         X2              =   11400
         Y1              =   450
         Y2              =   7680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   21
         X1              =   8160
         X2              =   8160
         Y1              =   450
         Y2              =   7680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   22
         X1              =   6105
         X2              =   6105
         Y1              =   450
         Y2              =   7680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   23
         X1              =   4080
         X2              =   4080
         Y1              =   450
         Y2              =   7680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   24
         X1              =   2055
         X2              =   2055
         Y1              =   450
         Y2              =   7680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -63600
         X2              =   -63600
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -66840
         X2              =   -66840
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   -68895
         X2              =   -68895
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   -70920
         X2              =   -70920
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   -72945
         X2              =   -72945
         Y1              =   450
         Y2              =   8070
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   -75000
         X2              =   -59820
         Y1              =   2490
         Y2              =   2490
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   -75000
         X2              =   -59820
         Y1              =   5970
         Y2              =   5970
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   0
         X2              =   15150
         Y1              =   7680
         Y2              =   7680
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   2
      Left            =   11190
      Top             =   555
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "取试料重量"
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
   Begin CSTextLibCtl.sidbEdit sdb_SMP_STD_WGT 
      Height          =   310
      Left            =   12540
      TabIndex        =   251
      Top             =   555
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   547
      _StockProps     =   125
      Text            =   " 0.00"
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
      NumIntDigits    =   5
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
End
Attribute VB_Name = "AQA0100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      客户特殊要求材质输入
'-- Program ID        AQA0100C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       客户特殊要求材质输入
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_CUST_SPEC_NO, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_SMP_LOC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_SMP_LOC_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_SMP_LEN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_PRE_SMP_QTY, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_SMP_STD_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          Call Gp_Ms_Collection(txt_TENCIL_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_YP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_YP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_YP_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_TS_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_TS_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_TS_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_RA_DIR_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_RA_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_RA_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_RA_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_EL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_EL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_EL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_YR_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_YR_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_YR_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_SNPP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_SNPP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_SNPP_EL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_SNPP_EL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_SNPP_EL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SG_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SG_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_SG_EL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_SG_EL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_SG_EL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_SP_EL_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_SP_EL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_SP_EL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_SP_EL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_TENCIL_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_HGT_TENCIL_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_HGT_TENCIL_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_YP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_YP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HGT_YP_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_TS_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_TS_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HGT_TS_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_RA_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_RA_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HGT_RA_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_HGT_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_HGT_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_EL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_HGT_EL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HGT_EL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_SNPP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_SNPP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_HGT_SNPP_EL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_SP_EL_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HGT_SP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HGT_SP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_SP_EL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_SP_EL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_SP_EL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_IMPACT_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
          Call Gp_Ms_Collection(txt_IMPACT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_A_IMPACT_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
        Call Gp_Ms_Collection(txt_A_IMPACT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TIM_IMPACT_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
      Call Gp_Ms_Collection(txt_TIM_IMPACT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_HARD_TYP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_HARD_TYP(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_HARD_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_HARD_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_HARD_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_BEND_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_BEND_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_BEND_ANGLE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_BEND_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
            Call Gp_Ms_Collection(txt_DWTT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
                Call Gp_Ms_Collection(txt_INS_EMP, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
   
    'MASTER Collection
    Mc1.Add Item:="AQA0101C.P_REFER", Key:="P-R"
    Mc1.Add Item:="AQA0101C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AQA0101C.P_ONEROW", Key:="P-O"
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
         
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_CUST_SPEC_NO"                  '客户特殊要求编号
            sCode = "CUST_SPEC_NO"
        
        Case "txt_SMP_LOC"                  '取试料位置
            sCode = "Q0042"
            Set oCodeName = txt_SMP_LOC_NAME
            
        '取样代码
        Case "txt_TENCIL_SMP_CD", "txt_SP_EL_SMP_CD", "txt_HGT_TENCIL_SMP_CD", "txt_HGT_SP_EL_SMP_CD", "txt_IMPACT_SMP_CD", "txt_TIM_IMPACT_SMP_CD", "txt_BEND_SMP_CD", "txt_RPT_BEND_SMP_CD", "txt_FOAT_SMP_CD", "txt_JOMINY_SMP_CD", "txt_HIC_SMP_CD", "txt_SSCC_SMP_CD", "txt_DWTT_SMP_CD", "txt_RMV_CAR_SMP_CD", "txt_FRACT_SMP_CD", "txt_NON_METAL_SMP_CD", "txt_A_IMPACT_SMP_CD", "txt_A_TIM_IMPACT_SMP_CD"
            If KeyCode = vbKeyF4 Then Call subSampCdPopup
            Exit Sub
                
        '判定
        Case "txt_YP_DSC_CD", "txt_TS_DSC_CD", "txt_RA_DSC_CD", "txt_EL_DSC_CD", "txt_YR_DSC_CD", "txt_SNPP_EL_DSC_CD", "txt_SG_EL_DSC_CD", "txt_SP_EL_DSC_CD", "txt_HGT_YP_DSC_CD", "txt_HGT_TS_DSC_CD", "txt_HGT_RA_DSC_CD", "txt_HGT_EL_DSC_CD", "txt_HGT_SNPP_EL_DSC_CD", "txt_HGT_SP_EL_DSC_CD", "txt_IMPACT_DSC_CD", "txt_TIM_IMPACT_DSC_CD", "txt_HARD_DSC_CD", "txt_BEND_DSC_CD", "txt_RPT_BEND_DSC_CD", "txt_WLD_HARD_DSC_CD", "txt_WLD_BEND_DSC_CD", "txt_UST_DSC_CD", "txt_FOAT_DSC_CD", "txt_JOMINY_DSC_CD", "txt_HIC_DSC_CD", "txt_SSCC_DSC_CD", "txt_DWTT_DSC_CD", "txt_RMV_CAR_DSC_CD", "txt_GRAIN_SIZE_DSC_CD", "txt_S_PRINT_DSC_CD", "txt_ACD_DSC_CD", "txt_FRACT_DSC_CD", "txt_NON_METAL_DSC_CD", "txt_BELT_STR_DSC_CD", "txt_A_IMPACT_DSC_CD", "txt_A_TIM_IMPACT_DSC_CD"
            sCode = "Q0002"
                                
        '试验温度
        Case "txt_HGT_TENCIL_TMP_UNIT", "txt_IMPACT_TMP_UNIT", "txt_TIM_IMPACT_TMP_UNIT", "txt_DWTT_TMP_UNIT", "txt_GRAIN_SIZE_TMP_UNIT", "txt_A_IMPACT_TMP_UNIT", "txt_A_TIM_IMPACT_TMP_UNIT"
            sCode = "Q0003"
        Case "txt_RA_DIR_CD"
            sCode = "Q0058"
            Set oCodeName = txt_RA_DIR_NAME
                                                 
        Case "txt_EL_CD"                    '断后伸长率
            sCode = "Q0004"
            Set oCodeName = txt_EL_CD(1)
            
        Case "txt_SNPP_EL_CD"               '规定非比例伸长应力
            sCode = "Q0005"
            Set oCodeName = txt_SNPP_EL_CD(1)
                    
        Case "txt_SG_EL_CD"                 '规定残余伸长应力
            sCode = "Q0006"
            Set oCodeName = txt_SG_EL_CD(1)
                         
        Case "txt_SP_EL_CD"                 '规定残余伸长应力
            sCode = "Q0007"
            Set oCodeName = txt_SP_EL_CD(1)
        
        Case "txt_HGT_EL_CD"                '高温拉伸试验 - 断后伸长率
            sCode = "Q0004"
            Set oCodeName = txt_HGT_EL_CD(1)
            
        Case "txt_HGT_SNPP_EL_CD"           '高温拉伸试验 - 规定非比例伸长应力
            sCode = "Q0005"
            Set oCodeName = txt_HGT_SNPP_EL_CD(1)
                    
        Case "txt_HGT_SP_EL_CD"             '高温拉伸试验 - 规定残余伸长应力
            sCode = "Q0007"
            Set oCodeName = txt_HGT_SP_EL_CD(1)
            
        Case "txt_IMPACT_KND"               '冲击试验 - 其它代码
            sCode = "Q0008"
            Set oCodeName = txt_IMPACT_KND(1)
                        
        Case "txt_IMPACT_DIR"            '冲击试验 - 其它代码
            sCode = "Q0009"
            Set oCodeName = txt_IMPACT_DIR(1)
            
        Case "txt_A_IMPACT_KND"             '追加冲击试验 - 其它代码
            sCode = "Q0008"
            Set oCodeName = txt_A_IMPACT_KND(1)
                        
        Case "txt_A_IMPACT_DIR"          '追加冲击试验 - 其它代码
            sCode = "Q0009"
            Set oCodeName = txt_A_IMPACT_DIR(1)
                        
        Case "txt_TIM_IMPACT_KND"           '时效冲击试验 - 其它代码
            sCode = "Q0008"
            Set oCodeName = txt_TIM_IMPACT_KND(1)
            
        Case "txt_TIM_IMPACT_DIR"        '时效冲击试验 - 其它代码
            sCode = "Q0009"
            Set oCodeName = txt_TIM_IMPACT_DIR(1)
                        
        Case "txt_A_TIM_IMPACT_KND"         '追加时效冲击试验 - 其它代码
            sCode = "Q0008"
            Set oCodeName = txt_A_TIM_IMPACT_KND(1)
            
        Case "txt_A_TIM_IMPACT_DIR"      '追加时效冲击试验 - 其它代码
            sCode = "Q0009"
            Set oCodeName = txt_A_TIM_IMPACT_DIR(1)
                        
        Case "txt_HARD_TYP"                 '硬度
            sCode = "Q0010"
            Set oCodeName = txt_HARD_TYP(1)
                    
        Case "txt_WLD_HARD_TYP"             '焊缝硬度
            sCode = "Q0011"
            Set oCodeName = txt_WLD_HARD_TYP(1)
            
        Case "txt_UST_STD_CD"               '超声波探伤（UST）
            sCode = "Q0046"
            Set oCodeName = txt_UST_STD_CD(1)
                        
        Case "txt_JOMINY_TYP"               '淬透性
            sCode = "Q0012"
            Set oCodeName = txt_JOMINY_TYP(1)
            
        Case "txt_HIC_SVT_KND"              '抗氢裂能力
            sCode = "Q0013"
            Set oCodeName = txt_HIC_SVT_KND(1)
            
        Case "txt_SSCC_SVT_KND"             '硫化物腐蚀裂纹
            sCode = "Q0014"
            Set oCodeName = txt_SSCC_SVT_KND(1)
            
        Case "txt_UST_GRD"                  'UST GRD
            sCode = "Q0053"
            Set oCodeName = txt_UST_GRD
            Set oCodeName = txt_UST_GRD_NAME
            
        Case "txt_RMV_CAR_TYP"              '脱碳层
            sCode = "Q0015"
            Set oCodeName = txt_RMV_CAR_TYP(1)
            
        Case "txt_GRAIN_SIZE_MTH"           '晶粒度
            sCode = "Q0016"
            Set oCodeName = txt_GRAIN_SIZE_MTH(1)
                        
        Case "txt_NON_METAL_TYP"            '非金属夹杂
            sCode = "Q0018"
            Set oCodeName = txt_NON_METAL_TYP(1)
                        
        Case "txt_NON_METAL_ACD1"           '非金属夹杂 - 粗系 - 1
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD1(1)
            
        Case "txt_NON_METAL_ACD2"           '非金属夹杂 - 粗系 - 2
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD2(1)
            
        Case "txt_NON_METAL_ACD3"           '非金属夹杂 - 粗系 - 3
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD3(1)
            
        Case "txt_NON_METAL_ACD4"           '非金属夹杂 - 粗系 - 4
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD4(1)
            
        Case "txt_NON_METAL_BCD1"           '非金属夹杂 - 细系 - 1
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD1(1)
            
        Case "txt_NON_METAL_BCD2"           '非金属夹杂 - 细系 - 2
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD2(1)
            
        Case "txt_NON_METAL_BCD3"           '非金属夹杂 - 细系 - 3
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD3(1)
            
        Case "txt_NON_METAL_BCD4"           '非金属夹杂 - 细系 - 4
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD4(1)
            
                        
        Case "txt_ACD_DFT_TYP1"             '酸浸检验 非金属夹杂 - 1
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP1(1)
            
        Case "txt_ACD_DFT_TYP2"             '非金属夹杂 - 2
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP2(1)
            
        Case "txt_ACD_DFT_TYP3"             '非金属夹杂 - 3
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP3(1)
            
        Case "txt_ACD_DFT_TYP4"             '非金属夹杂 - 4
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP4(1)
            
        Case "txt_ACD_DFT_TYP5"             '非金属夹杂 - 5
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP5(1)
                                    
        Case "txt_FRACT_NAME_CD1"           '断口检验 - 1
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD1(1)
                        
        Case "txt_FRACT_NAME_CD2"           '断口检验 - 2
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD2(1)
            
        Case "txt_FRACT_NAME_CD3"           '断口检验 - 3
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD3(1)
            
        Case "txt_FRACT_NAME_CD4"           '断口检验 - 4
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD4(1)
            
        Case "txt_FRACT_NAME_CD5"           '断口检验 - 5
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD5(1)
            
        Case "txt_BELT_STR_GRD"             '带状组织
            sCode = "Q0055"
            
        Case Else
            Exit Sub
        
    End Select
    
    If sCode = "" Then Exit Sub
        
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    If Len(Trim(txt_CUST_SPEC_NO.Text)) <> 0 Then
        Call Form_Ref
    End If


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
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_GetSampleCode           '取样代码检索
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Call GP_BACKCOLOR_WHITE(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    pControl(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()
        
    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call GP_BACKCOLOR_WHITE(Mc1("rControl"))
        Call Gp_Ms_NeceColor(Mc1("nControl"))
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If

    
End Sub

Public Sub Form_Pro()
       
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        If funFormCheck = False Then Exit Sub
        If Not (GF_Necessary_Value_Check(txt_SMP_LOC, "取试料位置", txt_SMP_LOC)) Then Exit Sub
        If Not (GF_Necessary_Value_Check(txt_SMP_LEN, "取试料长度", txt_SMP_LEN)) Then Exit Sub
 '       If Not (GF_Necessary_Value_Check(txt_PRE_SMP_QTY, "复样数量", txt_PRE_SMP_QTY)) Then Exit Sub
 '       If Not (GF_Necessary_Value_Check(sdb_SMP_STD_WGT, "取试料重量", sdb_SMP_STD_WGT)) Then Exit Sub
        txt_INS_EMP.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub


'---------------------------------------------------------------------------------
'---------------------------- Input Check ----------------------------------------
'---------------------------------------------------------------------------------
Private Function funFormCheck() As Boolean

    Dim iCnt As Integer

'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 1 ( 拉伸试验 ) -------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------

'屈服强度
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_YP_MIN, sdb_YP_MAX, txt_TENCIL_SMP_CD, txt_YP_DSC_CD) = False Then iCnt = iCnt + 1

'抗拉强度
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_TS_MIN, sdb_TS_MAX, txt_TS_DSC_CD) = False Then iCnt = iCnt + 1

'断面收缩率
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_RA_MIN, sdb_RA_MAX, txt_RA_DSC_CD) = False Then iCnt = iCnt + 1

'断后伸长率
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_EL_MIN, sdb_EL_MAX, txt_EL_CD(0), txt_EL_DSC_CD) = False Then iCnt = iCnt + 1
    
'屈强比
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_YR_MIN, sdb_YR_MAX, txt_YR_DSC_CD) = False Then iCnt = iCnt + 1
    
'规定非比例伸长应力
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_SNPP_EL_MIN, sdb_SNPP_EL_MAX, txt_SNPP_EL_CD(0), txt_SNPP_EL_DSC_CD) = False Then iCnt = iCnt + 1

'规定总伸长应力
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_SG_EL_MIN, sdb_SG_EL_MAX, txt_SG_EL_CD(0), txt_SG_EL_DSC_CD) = False Then iCnt = iCnt + 1

'规定残余伸长应力
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_SP_EL_MIN, sdb_SP_EL_MAX, txt_SP_EL_SMP_CD, txt_SP_EL_CD(0), txt_SP_EL_DSC_CD) = False Then iCnt = iCnt + 1
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 2 ( 高温拉伸试验 ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
'屈服强度
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_YP_MIN, sdb_HGT_YP_MAX, txt_HGT_TENCIL_SMP_CD, txt_HGT_TENCIL_TMP, txt_HGT_TENCIL_TMP_UNIT, txt_HGT_YP_DSC_CD) = False Then iCnt = iCnt + 1

'抗拉强度
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_TS_MIN, sdb_HGT_TS_MAX, txt_HGT_TS_DSC_CD) = False Then iCnt = iCnt + 1

'断面收缩率
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_RA_MIN, sdb_HGT_RA_MAX, txt_HGT_RA_DSC_CD) = False Then iCnt = iCnt + 1

'断后伸长率
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_EL_MIN, sdb_HGT_EL_MAX, txt_HGT_EL_CD(0), txt_HGT_EL_DSC_CD) = False Then iCnt = iCnt + 1
    
'规定非比例伸长应力
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_SNPP_EL_MIN, sdb_HGT_SNPP_EL_MAX, txt_HGT_SNPP_EL_CD(0), txt_HGT_SNPP_EL_DSC_CD) = False Then iCnt = iCnt + 1

'规定残余伸长应力
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_SP_EL_MIN, sdb_HGT_SP_EL_MAX, txt_HGT_SP_EL_SMP_CD, txt_HGT_SP_EL_CD(0), txt_HGT_SP_EL_DSC_CD) = False Then iCnt = iCnt + 1
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 3 ( 冲击、时效 ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------

'冲击试验
    If GF_MATR_IMPACT_INPUT_CHECK(txt_IMPACT_SMP_CD, txt_IMPACT_KND(0), txt_IMPACT_DIR(0), sdb_IMPACT_MIN, sdb_IMPACT_MIN_MIN, sdb_IMPACT_AVE_MIN, sdb_IMPACT_RATE_MIN, sdb_IMPACT_RATE_MAX, txt_IMPACT_DSC_CD) = False Then iCnt = iCnt + 1

'追加冲击试验
    If GF_MATR_IMPACT_INPUT_CHECK(txt_A_IMPACT_SMP_CD, txt_A_IMPACT_KND(0), txt_A_IMPACT_DIR(0), sdb_A_IMPACT_MIN, sdb_A_IMPACT_MIN_MIN, sdb_A_IMPACT_AVE_MIN, sdb_A_IMPACT_RATE_MIN, sdb_A_IMPACT_RATE_MAX, txt_A_IMPACT_DSC_CD) = False Then iCnt = iCnt + 1

'时效冲击试验
    If GF_MATR_TIM_IMPACT_INPUT_CHECK(txt_TIM_IMPACT_SMP_CD, txt_TIM_IMPACT_KND(0), txt_TIM_IMPACT_DIR(0), sdb_TIM_IMPACT_MIN, sdb_TIM_IMPACT_MIN_MIN, sdb_TIM_IMPACT_AVE_MIN, sdb_TIM_IMPACT_RATE_MIN, sdb_TIM_IMPACT_RATE_MAX, txt_TIM_IMPACT_DSC_CD) = False Then iCnt = iCnt + 1
    
'追加时效冲击试验
    If GF_MATR_TIM_IMPACT_INPUT_CHECK(txt_A_TIM_IMPACT_SMP_CD, txt_A_TIM_IMPACT_KND(0), txt_A_TIM_IMPACT_DIR(0), sdb_A_TIM_IMPACT_MIN, sdb_A_TIM_IMPACT_MIN_MIN, sdb_A_TIM_IMPACT_AVE_MIN, sdb_A_TIM_IMPACT_RATE_MIN, sdb_A_TIM_IMPACT_RATE_MAX, txt_A_TIM_IMPACT_DSC_CD) = False Then iCnt = iCnt + 1
    
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 4 ( 其它 ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
'硬度
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HARD_MIN, sdb_HARD_MAX, txt_HARD_TYP(0), txt_HARD_DSC_CD) = False Then iCnt = iCnt + 1
    
'弯曲试验
    If GF_MATR_COMMON_INPUT_CHECK(txt_BEND_SMP_CD, sdb_BEND_DIA, sdb_BEND_ANGLE, txt_BEND_DSC_CD) = False Then iCnt = iCnt + 1
    
'反复弯曲
    If GF_MATR_COMMON_INPUT_CHECK(txt_RPT_BEND_SMP_CD, sdb_RPT_BEND_TMS, txt_RPT_BEND_DSC_CD) = False Then iCnt = iCnt + 1

'焊缝硬度
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_WLD_HARD_MIN, sdb_WLD_HARD_MAX, txt_WLD_HARD_TYP(0), txt_WLD_HARD_UNIT, txt_WLD_HARD_DSC_CD) = False Then iCnt = iCnt + 1

'焊缝弯曲
    If GF_MATR_COMMON_INPUT_CHECK(sdb_WLD_BEND_DIA, sdb_WLD_BEND_ANG, txt_WLD_BEND_DSC_CD) = False Then iCnt = iCnt + 1

'超声波探伤（UST）
    If GF_MATR_COMMON_INPUT_CHECK(txt_UST_STD_CD(0), txt_UST_GRD, txt_UST_DSC_CD) = False Then iCnt = iCnt + 1

'锻平
    If GF_MATR_COMMON_INPUT_CHECK(txt_FOAT_SMP_CD, txt_FOAT_DSC_CD) = False Then iCnt = iCnt + 1

'淬透性
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_JOMINY_MIN, sdb_JOMINY_MAX, txt_JOMINY_SMP_CD, txt_JOMINY_TYP(0), sdb_JOMINY_DIST, txt_JOMINY_DSC_CD) = False Then iCnt = iCnt + 1

'抗氢裂能力
    If GF_MATR_HIC_INPUT_CHECK(txt_HIC_SMP_CD, txt_HIC_SVT_KND(0), sdb_HIC_CSR_MAX, sdb_HIC_CLR_MAX, sdb_HIC_CWR_MAX, txt_HIC_DSC_CD) = False Then iCnt = iCnt + 1

'硫化物腐蚀裂纹
    If GF_MATR_COMMON_INPUT_CHECK(txt_SSCC_SMP_CD, txt_SSCC_SVT_KND(0), sdb_SSCC_YP_TIM, sdb_SSCC_YP_MAX, txt_SSCC_DSC_CD) = False Then iCnt = iCnt + 1

'重力撕裂试验
    If GF_MATR_COMMON_INPUT_CHECK(txt_DWTT_SMP_CD, txt_DWTT_TMP, txt_DWTT_TMP_UNIT, sdb_DWTT_YP_MIN, txt_DWTT_DSC_CD, sdb_DWTT_YP_AVE) = False Then iCnt = iCnt + 1


'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 5 ( 金相检验 ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
'脱碳层
    If GF_MATR_COMMON_INPUT_CHECK(txt_RMV_CAR_SMP_CD, txt_RMV_CAR_TYP(0), sdb_RMV_CAR_MAX, txt_RMV_CAR_DSC_CD) = False Then iCnt = iCnt + 1
    
'晶粒度
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_GRAIN_SIZE_MIN, sdb_GRAIN_SIZE_MAX, txt_GRAIN_SIZE_MTH(0), sdb_GRAIN_SIZE_TMP, txt_GRAIN_SIZE_TMP_UNIT, sdb_GRAIN_SIZE_TIME, txt_GRAIN_SIZE_DSC_CD) = False Then iCnt = iCnt + 1
    
'硫印
    If GF_MATR_COMMON_INPUT_CHECK(sdb_S_PRINT_DRG, txt_S_PRINT_DSC_CD) = False Then iCnt = iCnt + 1

'酸浸检验
    If GF_MATR_ACD_DFT_INPUT_CHECK(txt_ACD_DFT_TYP1(0), sdb_ACD_DFT_GRD1, txt_ACD_DFT_TYP2(0), sdb_ACD_DFT_GRD2, txt_ACD_DFT_TYP3(0), sdb_ACD_DFT_GRD3, txt_ACD_DFT_TYP4(0), sdb_ACD_DFT_GRD4, txt_ACD_DFT_TYP5(0), sdb_ACD_DFT_GRD5, txt_ACD_DSC_CD) = False Then iCnt = iCnt + 1

'断口检验
    If GF_MATR_FRACT_INPUT_CHECK(txt_FRACT_SMP_CD, txt_FRACT_NAME_CD1(0), txt_FRACT_GRD1, txt_FRACT_NAME_CD2(0), txt_FRACT_GRD2, txt_FRACT_NAME_CD3(0), txt_FRACT_GRD3, txt_FRACT_NAME_CD4(0), txt_FRACT_GRD4, txt_FRACT_NAME_CD5(0), txt_FRACT_GRD5, txt_FRACT_DSC_CD) = False Then iCnt = iCnt + 1
    
'非金属夹杂
    If GF_MATR_NON_METAL_INPUT_CHECK(txt_NON_METAL_SMP_CD, txt_NON_METAL_TYP(0), txt_NON_METAL_ACD1(0), sdb_NON_METAL_AGRD1, txt_NON_METAL_ACD2(0), sdb_NON_METAL_AGRD2, txt_NON_METAL_ACD3(0), sdb_NON_METAL_AGRD3, txt_NON_METAL_ACD4(0), sdb_NON_METAL_AGRD4, txt_NON_METAL_BCD1(0), sdb_NON_METAL_BGRD1, txt_NON_METAL_BCD2(0), sdb_NON_METAL_BGRD2, txt_NON_METAL_BCD3(0), sdb_NON_METAL_BGRD3, txt_NON_METAL_BCD4(0), sdb_NON_METAL_BGRD4, txt_NON_METAL_DSC_CD) = False Then iCnt = iCnt + 1
    
    If iCnt = 0 Then
        funFormCheck = True
    Else
        Call Gp_MsgBoxDisplay("Correct error field")
    End If
    
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------- 取样代码 -----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub subSampCdPopup()
    sSampSearch = Me.ActiveControl.Text
    frmSampStd.Show 1
    If sSampCd <> "" Then Me.ActiveControl.Text = sSampCd
End Sub


