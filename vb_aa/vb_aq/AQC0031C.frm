VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQC0031C 
   Caption         =   "材质试验实绩输入 - AQC0031C"
   ClientHeight    =   9045
   ClientLeft      =   15
   ClientTop       =   1710
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_ins_emp 
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
      Left            =   10140
      TabIndex        =   211
      Tag             =   "INS_EMP"
      Top             =   90
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txt_knd 
      Height          =   315
      Left            =   10830
      TabIndex        =   210
      Tag             =   "99"
      Top             =   90
      Visible         =   0   'False
      Width           =   465
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   465
      Left            =   6540
      TabIndex        =   208
      Top             =   60
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   820
      _Version        =   196609
      Begin Threed.SSRibbon sbtn_SMP_TYPE_SELECT 
         Height          =   375
         Left            =   30
         TabIndex        =   209
         Top             =   30
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "录入状态：常规样录入"
      End
   End
   Begin VB.TextBox txt_SMP_CUT_LOC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5790
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   119
      Tag             =   "99"
      Top             =   120
      Width           =   435
   End
   Begin VB.TextBox txt_SMP_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1620
      MaxLength       =   14
      TabIndex        =   118
      Tag             =   "99"
      Top             =   120
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8100
      Left            =   90
      TabIndex        =   138
      Top             =   1245
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14288
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "拉伸／高温拉伸／其它"
      TabPicture(0)   =   "AQC0031C.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line2(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line3(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line4(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line4(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line4(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line5(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line5(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line5(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line5(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line5(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line5(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line5(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line5(7)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line5(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line5(9)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line5(10)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line6(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line5(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line6(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line5(12)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line5(13)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(4)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(5)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Line5(15)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "sdb_RA_RST_AVE"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "sdb_RA_RST_3"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "sdb_RA_RST_2"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "sdb_YR_RST"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "sdb_HGT_SP_EL_RST"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "sdb_HGT_SNPP_EL_RST"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "sdb_HGT_EL_RST"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "sdb_HGT_RA_RST"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "sdb_HGT_TS_RST"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "sdb_HGT_YP_RST"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "sdb_SG_EL_RST"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "sdb_SP_EL_RST"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "sdb_SNPP_EL_RST"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "sdb_EL_RST"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "sdb_RA_RST_1"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "sdb_TS_RST"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "sdb_YP_RST"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "sdb_DWTT_YP_RST3"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "sdb_DWTT_YP_RST2"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "sivbLB1(29)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "sdb_DWTT_YP_RST1"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "sdb_SSCC_YP_RST"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "sdb_HIC_CWR_RST"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "sdb_HIC_CLR_RST"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "sdb_HIC_CSR_RST"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "sdb_WLD_HARD_RST"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "sdb_HARD_RST"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "sivbLB1(5)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "sivbLB1(4)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "sivbLB1(3)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "sivbLB1(2)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "sivbLB1(1)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "sivbLB1(0)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "ULabel1(9)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "sivbLB1(24)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "sivbLB1(23)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "sivbLB1(22)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "sivbLB1(21)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "sivbLB1(20)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "ULabel1(8)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "sivbLB1(19)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "sivbLB1(18)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "sivbLB1(17)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "sivbLB1(16)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "sivbLB1(15)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "sivbLB1(13)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "sivbLB1(11)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "sivbLB1(10)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "sivbLB1(9)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "sivbLB1(8)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "sivbLB1(7)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "sivbLB1(6)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "txt_HARD_TYP"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txt_WLD_HARD_TYP"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "txt_RPT_BEND_RST"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "cbo_BEND_RST"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "cbo_WLD_BEND_RST"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "cbo_FOAT_RST"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txt_BEND_RST"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "txt_WLD_BEND_RST"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "txt_FOAT_RST"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "txt_HARD_NAME"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "txt_WLD_HARD_NAME"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "sivbLB2"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).ControlCount=   94
      TabCaption(1)   =   "冲击／时效冲击"
      TabPicture(1)   =   "AQC0031C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line5(14)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Shape1(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Shape3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ULabel1(65)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ULabel1(64)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ULabel1(63)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "ULabel1(62)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "sdb_A_IMPACT_RATE_AVE_RST"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "sdb_A_IMPACT_RATE_RST6"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "sdb_A_IMPACT_RATE_RST5"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "sdb_A_IMPACT_RATE_RST4"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "sdb_A_IMPACT_RATE_RST3"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "sdb_A_IMPACT_RATE_RST2"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "sdb_A_IMPACT_RATE_RST1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "sdb_IMPACT_RATE_AVE_RST"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "sdb_IMPACT_RATE_RST6"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "sdb_IMPACT_RATE_RST5"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "sdb_IMPACT_RATE_RST4"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "sdb_IMPACT_RATE_RST3"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "sdb_IMPACT_RATE_RST2"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "sdb_IMPACT_RATE_RST1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "ULabel1(56)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "ULabel1(55)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "ULabel1(54)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "ULabel1(53)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "ULabel1(52)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "ULabel1(51)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "ULabel1(50)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "ULabel1(49)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "ULabel1(48)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "ULabel1(47)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "ULabel1(46)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "ULabel4(5)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "sdb_A_TIM_IMPACT_RATE_RST"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "sdb_A_TIM_IMPACT_RST_AVE"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "sdb_A_TIM_IMPACT_RST6"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "sdb_A_TIM_IMPACT_RST5"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "sdb_A_TIM_IMPACT_RST4"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "sdb_A_TIM_IMPACT_RST3"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "sdb_A_TIM_IMPACT_RST2"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "sdb_A_TIM_IMPACT_RST1"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "ULabel1(45)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "ULabel1(44)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "ULabel1(43)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "ULabel1(42)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "ULabel1(41)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "ULabel1(40)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "ULabel1(39)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "ULabel1(38)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "ULabel1(37)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "ULabel1(36)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "ULabel1(35)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "sdb_A_IMPACT_RST_AVE"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "sdb_A_IMPACT_RST6"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "sdb_A_IMPACT_RST5"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "sdb_A_IMPACT_RST4"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "sdb_A_IMPACT_RST3"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "sdb_A_IMPACT_RST2"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "sdb_A_IMPACT_RST1"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "ULabel1(34)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "ULabel1(33)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "ULabel1(32)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "ULabel1(31)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "ULabel1(30)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "ULabel1(29)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "ULabel1(28)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "ULabel1(27)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "ULabel1(26)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "ULabel1(25)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "ULabel1(24)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "ULabel4(4)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "ULabel1(23)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "ULabel1(22)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "ULabel1(21)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "ULabel1(20)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "ULabel1(19)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "ULabel1(18)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "ULabel1(17)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "ULabel1(16)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "ULabel1(15)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "ULabel1(14)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "ULabel1(13)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "ULabel4(2)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "ULabel4(1)"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "sdb_TIM_IMPACT_RATE_RST"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "sdb_TIM_IMPACT_RST_AVE"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "sdb_TIM_IMPACT_RST6"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "sdb_TIM_IMPACT_RST5"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "sdb_TIM_IMPACT_RST4"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "sdb_TIM_IMPACT_RST3"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "sdb_TIM_IMPACT_RST2"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).Control(94)=   "sdb_TIM_IMPACT_RST1"
      Tab(1).Control(94).Enabled=   0   'False
      Tab(1).Control(95)=   "sdb_IMPACT_RST_AVE"
      Tab(1).Control(95).Enabled=   0   'False
      Tab(1).Control(96)=   "sdb_IMPACT_RST6"
      Tab(1).Control(96).Enabled=   0   'False
      Tab(1).Control(97)=   "sdb_IMPACT_RST5"
      Tab(1).Control(97).Enabled=   0   'False
      Tab(1).Control(98)=   "sdb_IMPACT_RST4"
      Tab(1).Control(98).Enabled=   0   'False
      Tab(1).Control(99)=   "sdb_IMPACT_RST3"
      Tab(1).Control(99).Enabled=   0   'False
      Tab(1).Control(100)=   "sdb_IMPACT_RST2"
      Tab(1).Control(100).Enabled=   0   'False
      Tab(1).Control(101)=   "sdb_IMPACT_RST1"
      Tab(1).Control(101).Enabled=   0   'False
      Tab(1).Control(102)=   "txt_IMPACT_DIR"
      Tab(1).Control(102).Enabled=   0   'False
      Tab(1).Control(103)=   "txt_IMPACT_KND"
      Tab(1).Control(103).Enabled=   0   'False
      Tab(1).Control(104)=   "txt_IMPACT_KND_NAME"
      Tab(1).Control(104).Enabled=   0   'False
      Tab(1).Control(105)=   "txt_IMPACT_DIR_NAME"
      Tab(1).Control(105).Enabled=   0   'False
      Tab(1).Control(106)=   "txt_TIM_IMPACT_DIR_NAME"
      Tab(1).Control(106).Enabled=   0   'False
      Tab(1).Control(107)=   "txt_TIM_IMPACT_KND_NAME"
      Tab(1).Control(107).Enabled=   0   'False
      Tab(1).Control(108)=   "txt_TIM_IMPACT_KND"
      Tab(1).Control(108).Enabled=   0   'False
      Tab(1).Control(109)=   "txt_TIM_IMPACT_DIR"
      Tab(1).Control(109).Enabled=   0   'False
      Tab(1).Control(110)=   "txt_A_IMPACT_DIR_NAME"
      Tab(1).Control(110).Enabled=   0   'False
      Tab(1).Control(111)=   "txt_A_IMPACT_KND_NAME"
      Tab(1).Control(111).Enabled=   0   'False
      Tab(1).Control(112)=   "txt_A_IMPACT_KND"
      Tab(1).Control(112).Enabled=   0   'False
      Tab(1).Control(113)=   "txt_A_IMPACT_DIR"
      Tab(1).Control(113).Enabled=   0   'False
      Tab(1).Control(114)=   "txt_A_TIM_IMPACT_DIR"
      Tab(1).Control(114).Enabled=   0   'False
      Tab(1).Control(115)=   "txt_A_TIM_IMPACT_KND"
      Tab(1).Control(115).Enabled=   0   'False
      Tab(1).Control(116)=   "txt_A_TIM_IMPACT_KND_NAME"
      Tab(1).Control(116).Enabled=   0   'False
      Tab(1).Control(117)=   "txt_A_TIM_IMPACT_DIR_NAME"
      Tab(1).Control(117).Enabled=   0   'False
      Tab(1).Control(118)=   "Cob_IMPACT_SIZE"
      Tab(1).Control(118).Enabled=   0   'False
      Tab(1).Control(119)=   "Cob_TIM_IMPACT_SIZE"
      Tab(1).Control(119).Enabled=   0   'False
      Tab(1).Control(120)=   "Cob_A_IMPACT_SIZE"
      Tab(1).Control(120).Enabled=   0   'False
      Tab(1).Control(121)=   "Cob_A_TIM_IMPACT_SIZE"
      Tab(1).Control(121).Enabled=   0   'False
      Tab(1).Control(122)=   "TXT_IMPACT_SIZE_CD"
      Tab(1).Control(122).Enabled=   0   'False
      Tab(1).Control(123)=   "TXT_A_IMPACT_SIZE_CD"
      Tab(1).Control(123).Enabled=   0   'False
      Tab(1).Control(124)=   "TXT_TIM_IMPACT_SIZE_CD"
      Tab(1).Control(124).Enabled=   0   'False
      Tab(1).Control(125)=   "TXT_A_TIM_IMPACT_SIZE_CD"
      Tab(1).Control(125).Enabled=   0   'False
      Tab(1).ControlCount=   126
      TabCaption(2)   =   "金相检验"
      TabPicture(2)   =   "AQC0031C.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line3(24)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Line3(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Line49(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Line49(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Line49(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Line3(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Line3(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Line3(4)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Line3(5)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Line49(5)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Line3(6)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Line49(6)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Line3(7)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Line3(8)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Line3(9)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Line3(10)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Line3(11)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Line49(7)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Line49(1)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Line49(8)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Line49(9)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Line49(10)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Line49(11)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Line49(12)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Line49(13)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Line49(15)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Line49(16)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "sdb_TIN_GRD"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "ULabel87(3)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "sdb_DS_GRD"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "ULabel87(0)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "ULabel1(66)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "sdb_OST_GRAIN_SIZE_RST"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "txt_BELT_STR_GRD_RST"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "ULabel4(15)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "ULabel4(14)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "ULabel4(13)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "ULabel4(12)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "ULabel27(3)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "sdb_JOMINY_DIST3"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "sdb_JOMINY_DIST2"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "sdb_JOMINY_DIST1"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "sdb_JOMINY_RST_TOP3"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "sdb_JOMINY_RST_TOP2"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "sdb_JOMINY_RST_TOP1"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "sdb_FRACT_GRD_RST5"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "sdb_FRACT_GRD_RST4"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "sdb_FRACT_GRD_RST3"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "sdb_FRACT_GRD_RST2"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "sdb_FRACT_GRD_RST1"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "ULabel4(11)"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "ULabel4(10)"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "ULabel4(9)"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "ULabel27(25)"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "ULabel87(2)"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "ULabel87(1)"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "ULabel27(0)"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "ULabel4(8)"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "ULabel4(7)"
      Tab(2).Control(58).Enabled=   0   'False
      Tab(2).Control(59)=   "ULabel4(6)"
      Tab(2).Control(59).Enabled=   0   'False
      Tab(2).Control(60)=   "ULabel27(23)"
      Tab(2).Control(60).Enabled=   0   'False
      Tab(2).Control(61)=   "sdb_ACD_RST5"
      Tab(2).Control(61).Enabled=   0   'False
      Tab(2).Control(62)=   "sdb_ACD_RST4"
      Tab(2).Control(62).Enabled=   0   'False
      Tab(2).Control(63)=   "sdb_ACD_RST1"
      Tab(2).Control(63).Enabled=   0   'False
      Tab(2).Control(64)=   "sdb_ACD_RST3"
      Tab(2).Control(64).Enabled=   0   'False
      Tab(2).Control(65)=   "sdb_ACD_RST2"
      Tab(2).Control(65).Enabled=   0   'False
      Tab(2).Control(66)=   "ULabel1(60)"
      Tab(2).Control(66).Enabled=   0   'False
      Tab(2).Control(67)=   "ULabel1(59)"
      Tab(2).Control(67).Enabled=   0   'False
      Tab(2).Control(68)=   "ULabel1(58)"
      Tab(2).Control(68).Enabled=   0   'False
      Tab(2).Control(69)=   "ULabel4(34)"
      Tab(2).Control(69).Enabled=   0   'False
      Tab(2).Control(70)=   "ULabel4(3)"
      Tab(2).Control(70).Enabled=   0   'False
      Tab(2).Control(71)=   "ULabel1(57)"
      Tab(2).Control(71).Enabled=   0   'False
      Tab(2).Control(72)=   "ULabel4(0)"
      Tab(2).Control(72).Enabled=   0   'False
      Tab(2).Control(73)=   "sdb_NON_METAL_BRST1"
      Tab(2).Control(73).Enabled=   0   'False
      Tab(2).Control(74)=   "sdb_NON_METAL_BRST4"
      Tab(2).Control(74).Enabled=   0   'False
      Tab(2).Control(75)=   "sdb_NON_METAL_BRST3"
      Tab(2).Control(75).Enabled=   0   'False
      Tab(2).Control(76)=   "sdb_NON_METAL_ARST4"
      Tab(2).Control(76).Enabled=   0   'False
      Tab(2).Control(77)=   "sdb_NON_METAL_BRST2"
      Tab(2).Control(77).Enabled=   0   'False
      Tab(2).Control(78)=   "sdb_NON_METAL_ARST1"
      Tab(2).Control(78).Enabled=   0   'False
      Tab(2).Control(79)=   "sdb_NON_METAL_ARST2"
      Tab(2).Control(79).Enabled=   0   'False
      Tab(2).Control(80)=   "sdb_NON_METAL_ARST3"
      Tab(2).Control(80).Enabled=   0   'False
      Tab(2).Control(81)=   "sdb_S_PRINT_RST"
      Tab(2).Control(81).Enabled=   0   'False
      Tab(2).Control(82)=   "sdb_RMV_CAR_RST"
      Tab(2).Control(82).Enabled=   0   'False
      Tab(2).Control(83)=   "sdb_GRAIN_SIZE_RST"
      Tab(2).Control(83).Enabled=   0   'False
      Tab(2).Control(84)=   "txt_FRACT_NAME_CD4"
      Tab(2).Control(84).Enabled=   0   'False
      Tab(2).Control(85)=   "txt_FRACT_NAME_CD5"
      Tab(2).Control(85).Enabled=   0   'False
      Tab(2).Control(86)=   "txt_FRACT_NAME_CD4_NAME"
      Tab(2).Control(86).Enabled=   0   'False
      Tab(2).Control(87)=   "txt_FRACT_NAME_CD5_NAME"
      Tab(2).Control(87).Enabled=   0   'False
      Tab(2).Control(88)=   "txt_FRACT_NAME_CD3_NAME"
      Tab(2).Control(88).Enabled=   0   'False
      Tab(2).Control(89)=   "txt_FRACT_NAME_CD2_NAME"
      Tab(2).Control(89).Enabled=   0   'False
      Tab(2).Control(90)=   "txt_FRACT_NAME_CD1_NAME"
      Tab(2).Control(90).Enabled=   0   'False
      Tab(2).Control(91)=   "txt_FRACT_NAME_CD3"
      Tab(2).Control(91).Enabled=   0   'False
      Tab(2).Control(92)=   "txt_FRACT_NAME_CD2"
      Tab(2).Control(92).Enabled=   0   'False
      Tab(2).Control(93)=   "txt_FRACT_NAME_CD1"
      Tab(2).Control(93).Enabled=   0   'False
      Tab(2).Control(94)=   "txt_RMV_CAR_TYP_NAME"
      Tab(2).Control(94).Enabled=   0   'False
      Tab(2).Control(95)=   "txt_RMV_CAR_TYP"
      Tab(2).Control(95).Enabled=   0   'False
      Tab(2).Control(96)=   "txt_ACD_DFT_TYP4"
      Tab(2).Control(96).Enabled=   0   'False
      Tab(2).Control(97)=   "txt_ACD_DFT_TYP5"
      Tab(2).Control(97).Enabled=   0   'False
      Tab(2).Control(98)=   "txt_ACD_DFT_TYP4_NAME"
      Tab(2).Control(98).Enabled=   0   'False
      Tab(2).Control(99)=   "txt_ACD_DFT_TYP5_NAME"
      Tab(2).Control(99).Enabled=   0   'False
      Tab(2).Control(100)=   "txt_ACD_DFT_TYP3_NAME"
      Tab(2).Control(100).Enabled=   0   'False
      Tab(2).Control(101)=   "txt_ACD_DFT_TYP2_NAME"
      Tab(2).Control(101).Enabled=   0   'False
      Tab(2).Control(102)=   "txt_ACD_DFT_TYP1_NAME"
      Tab(2).Control(102).Enabled=   0   'False
      Tab(2).Control(103)=   "txt_ACD_DFT_TYP3"
      Tab(2).Control(103).Enabled=   0   'False
      Tab(2).Control(104)=   "txt_ACD_DFT_TYP2"
      Tab(2).Control(104).Enabled=   0   'False
      Tab(2).Control(105)=   "txt_ACD_DFT_TYP1"
      Tab(2).Control(105).Enabled=   0   'False
      Tab(2).Control(106)=   "txt_NON_METAL_BCD4"
      Tab(2).Control(106).Enabled=   0   'False
      Tab(2).Control(107)=   "txt_NON_METAL_BCD4_NAME"
      Tab(2).Control(107).Enabled=   0   'False
      Tab(2).Control(108)=   "txt_NON_METAL_BCD1"
      Tab(2).Control(108).Enabled=   0   'False
      Tab(2).Control(109)=   "txt_NON_METAL_BCD2"
      Tab(2).Control(109).Enabled=   0   'False
      Tab(2).Control(110)=   "txt_NON_METAL_BCD3"
      Tab(2).Control(110).Enabled=   0   'False
      Tab(2).Control(111)=   "txt_NON_METAL_BCD1_NAME"
      Tab(2).Control(111).Enabled=   0   'False
      Tab(2).Control(112)=   "txt_NON_METAL_BCD2_NAME"
      Tab(2).Control(112).Enabled=   0   'False
      Tab(2).Control(113)=   "txt_NON_METAL_BCD3_NAME"
      Tab(2).Control(113).Enabled=   0   'False
      Tab(2).Control(114)=   "txt_NON_METAL_ACD4"
      Tab(2).Control(114).Enabled=   0   'False
      Tab(2).Control(115)=   "txt_NON_METAL_ACD4_NAME"
      Tab(2).Control(115).Enabled=   0   'False
      Tab(2).Control(116)=   "txt_NON_METAL_ACD1"
      Tab(2).Control(116).Enabled=   0   'False
      Tab(2).Control(117)=   "txt_NON_METAL_ACD2"
      Tab(2).Control(117).Enabled=   0   'False
      Tab(2).Control(118)=   "txt_NON_METAL_ACD3"
      Tab(2).Control(118).Enabled=   0   'False
      Tab(2).Control(119)=   "txt_NON_METAL_ACD1_NAME"
      Tab(2).Control(119).Enabled=   0   'False
      Tab(2).Control(120)=   "txt_NON_METAL_ACD2_NAME"
      Tab(2).Control(120).Enabled=   0   'False
      Tab(2).Control(121)=   "txt_NON_METAL_ACD3_NAME"
      Tab(2).Control(121).Enabled=   0   'False
      Tab(2).Control(122)=   "ULabel27(1)"
      Tab(2).Control(122).Enabled=   0   'False
      Tab(2).Control(123)=   "ULabel27(2)"
      Tab(2).Control(123).Enabled=   0   'False
      Tab(2).Control(124)=   "txt_JOMINY_TYP"
      Tab(2).Control(124).Enabled=   0   'False
      Tab(2).Control(125)=   "txt_JOMINY_NAME"
      Tab(2).Control(125).Enabled=   0   'False
      Tab(2).Control(126)=   "Op_CHAGE"
      Tab(2).Control(126).Enabled=   0   'False
      Tab(2).Control(127)=   "Op_ONLY"
      Tab(2).Control(127).Enabled=   0   'False
      Tab(2).Control(128)=   "txt_SAVE_CASE"
      Tab(2).Control(128).Enabled=   0   'False
      Tab(2).ControlCount=   129
      Begin VB.TextBox txt_SAVE_CASE 
         Height          =   270
         Left            =   -71520
         TabIndex        =   215
         Text            =   "0"
         Top             =   450
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OptionButton Op_ONLY 
         Caption         =   "单独保存"
         Height          =   315
         Left            =   -73110
         TabIndex        =   214
         Top             =   450
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton Op_CHAGE 
         Caption         =   "按炉号保存"
         Height          =   315
         Left            =   -74520
         TabIndex        =   213
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox TXT_A_TIM_IMPACT_SIZE_CD 
         Height          =   315
         Left            =   -66210
         MaxLength       =   1
         TabIndex        =   70
         Tag             =   "27"
         Top             =   5160
         Width           =   405
      End
      Begin VB.TextBox TXT_TIM_IMPACT_SIZE_CD 
         Height          =   315
         Left            =   -66210
         MaxLength       =   1
         TabIndex        =   46
         Tag             =   "26"
         Top             =   1770
         Width           =   405
      End
      Begin VB.TextBox TXT_A_IMPACT_SIZE_CD 
         Height          =   315
         Left            =   -73620
         MaxLength       =   1
         TabIndex        =   55
         Tag             =   "25"
         Top             =   5160
         Width           =   405
      End
      Begin VB.TextBox TXT_IMPACT_SIZE_CD 
         Height          =   315
         Left            =   -73620
         MaxLength       =   1
         TabIndex        =   31
         Tag             =   "24"
         Top             =   1770
         Width           =   405
      End
      Begin VB.ComboBox Cob_A_TIM_IMPACT_SIZE 
         Height          =   300
         ItemData        =   "AQC0031C.frx":0054
         Left            =   -65790
         List            =   "AQC0031C.frx":0064
         Locked          =   -1  'True
         TabIndex        =   205
         Tag             =   "27"
         Top             =   5160
         Width           =   2145
      End
      Begin VB.ComboBox Cob_A_IMPACT_SIZE 
         Height          =   300
         ItemData        =   "AQC0031C.frx":0088
         Left            =   -73200
         List            =   "AQC0031C.frx":0098
         Locked          =   -1  'True
         TabIndex        =   203
         Tag             =   "25"
         Top             =   5160
         Width           =   2145
      End
      Begin VB.ComboBox Cob_TIM_IMPACT_SIZE 
         Height          =   300
         ItemData        =   "AQC0031C.frx":00BC
         Left            =   -65790
         List            =   "AQC0031C.frx":00CC
         Locked          =   -1  'True
         TabIndex        =   204
         Tag             =   "26"
         Top             =   1770
         Width           =   2145
      End
      Begin VB.ComboBox Cob_IMPACT_SIZE 
         Height          =   300
         ItemData        =   "AQC0031C.frx":00F0
         Left            =   -73170
         List            =   "AQC0031C.frx":0100
         Locked          =   -1  'True
         TabIndex        =   202
         Tag             =   "24"
         Top             =   1770
         Width           =   2145
      End
      Begin CSTextLibCtl.sivbLB sivbLB2 
         Height          =   2265
         Left            =   8130
         TabIndex        =   200
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "屈强比                                                                                              %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "屈强比                                                                                              %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin VB.TextBox txt_JOMINY_NAME 
         Height          =   300
         Left            =   -68400
         TabIndex        =   105
         Top             =   5100
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txt_JOMINY_TYP 
         Height          =   300
         Left            =   -68400
         TabIndex        =   102
         Tag             =   "20"
         Top             =   4725
         Width           =   735
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   2
         Left            =   -64530
         Top             =   3705
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   529
         Caption         =   "夹杂"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
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
         Index           =   1
         Left            =   -64500
         Top             =   4050
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   529
         Caption         =   "(级)"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
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
      Begin VB.TextBox txt_NON_METAL_ACD3_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   198
         Top             =   2370
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD2_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   197
         Top             =   1965
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD1_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   196
         Top             =   1500
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD3 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   131
         Tag             =   "34"
         Top             =   2370
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD2 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   130
         Tag             =   "34"
         Top             =   1965
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD1 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   129
         Tag             =   "34"
         Top             =   1500
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD4_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   195
         Top             =   2730
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD4 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   132
         Tag             =   "34"
         Top             =   2730
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD3_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   194
         Top             =   4140
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD2_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   193
         Top             =   3735
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD1_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   192
         Top             =   3330
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD3 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   135
         Tag             =   "34"
         Top             =   4140
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD2 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   134
         Tag             =   "34"
         Top             =   3735
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD1 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   133
         Tag             =   "34"
         Top             =   3330
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD4_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62700
         MaxLength       =   80
         TabIndex        =   191
         Top             =   4560
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD4 
         Height          =   315
         Left            =   -63150
         MaxLength       =   1
         TabIndex        =   136
         Tag             =   "34"
         Top             =   4560
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP1 
         Height          =   315
         Left            =   -68400
         MaxLength       =   2
         TabIndex        =   92
         Tag             =   "32"
         Top             =   1560
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP2 
         Height          =   315
         Left            =   -68400
         MaxLength       =   2
         TabIndex        =   94
         Tag             =   "32"
         Top             =   1995
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP3 
         Height          =   315
         Left            =   -68400
         MaxLength       =   2
         TabIndex        =   96
         Tag             =   "32"
         Top             =   2400
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP1_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67950
         MaxLength       =   80
         TabIndex        =   190
         Top             =   1560
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP2_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67950
         MaxLength       =   80
         TabIndex        =   189
         Top             =   1995
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP3_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67950
         MaxLength       =   80
         TabIndex        =   188
         Top             =   2400
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP5_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67950
         MaxLength       =   80
         TabIndex        =   187
         Top             =   3225
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP4_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67950
         MaxLength       =   80
         TabIndex        =   186
         Top             =   2790
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP5 
         Height          =   315
         Left            =   -68400
         MaxLength       =   2
         TabIndex        =   100
         Tag             =   "32"
         Top             =   3225
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP4 
         Height          =   315
         Left            =   -68400
         MaxLength       =   2
         TabIndex        =   98
         Tag             =   "32"
         Top             =   2790
         Width           =   435
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
         Left            =   -73230
         MaxLength       =   1
         TabIndex        =   128
         Tag             =   "29"
         Top             =   2610
         Width           =   495
      End
      Begin VB.TextBox txt_RMV_CAR_TYP_NAME 
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
         Left            =   -72720
         MaxLength       =   80
         TabIndex        =   185
         Top             =   2610
         Width           =   1470
      End
      Begin VB.TextBox txt_FRACT_NAME_CD1 
         Height          =   315
         Left            =   -73230
         MaxLength       =   2
         TabIndex        =   82
         Tag             =   "31"
         Top             =   3840
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD2 
         Height          =   315
         Left            =   -73230
         MaxLength       =   2
         TabIndex        =   84
         Tag             =   "31"
         Top             =   4245
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD3 
         Height          =   315
         Left            =   -73230
         MaxLength       =   2
         TabIndex        =   86
         Tag             =   "31"
         Top             =   4650
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD1_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72780
         MaxLength       =   80
         TabIndex        =   184
         Top             =   3840
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD2_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72780
         MaxLength       =   80
         TabIndex        =   183
         Top             =   4245
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD3_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72780
         MaxLength       =   80
         TabIndex        =   182
         Top             =   4650
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD5_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72780
         MaxLength       =   80
         TabIndex        =   181
         Top             =   5475
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD4_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72780
         MaxLength       =   80
         TabIndex        =   180
         Top             =   5070
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD5 
         Height          =   315
         Left            =   -73230
         MaxLength       =   2
         TabIndex        =   90
         Tag             =   "31"
         Top             =   5475
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD4 
         Height          =   315
         Left            =   -73230
         MaxLength       =   2
         TabIndex        =   88
         Tag             =   "31"
         Top             =   5070
         Width           =   435
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_DIR_NAME 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -62040
         TabIndex        =   179
         Top             =   4590
         Width           =   1695
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_KND_NAME 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65790
         TabIndex        =   178
         Top             =   4590
         Width           =   1695
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_KND 
         Height          =   315
         Left            =   -66210
         MaxLength       =   1
         TabIndex        =   126
         Tag             =   "26"
         Top             =   4590
         Width           =   405
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_DIR 
         Height          =   300
         Left            =   -62460
         MaxLength       =   1
         TabIndex        =   127
         Tag             =   "26"
         Top             =   4590
         Width           =   405
      End
      Begin VB.TextBox txt_A_IMPACT_DIR 
         Height          =   300
         Left            =   -69900
         MaxLength       =   1
         TabIndex        =   125
         Tag             =   "25"
         Top             =   4590
         Width           =   405
      End
      Begin VB.TextBox txt_A_IMPACT_KND 
         Height          =   300
         Left            =   -73620
         MaxLength       =   1
         TabIndex        =   124
         Tag             =   "25"
         Top             =   4590
         Width           =   405
      End
      Begin VB.TextBox txt_A_IMPACT_KND_NAME 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73200
         TabIndex        =   177
         Top             =   4590
         Width           =   1695
      End
      Begin VB.TextBox txt_A_IMPACT_DIR_NAME 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -69480
         TabIndex        =   176
         Top             =   4590
         Width           =   1695
      End
      Begin VB.TextBox txt_TIM_IMPACT_DIR 
         Height          =   315
         Left            =   -62460
         MaxLength       =   1
         TabIndex        =   123
         Tag             =   "26"
         Top             =   1140
         Width           =   405
      End
      Begin VB.TextBox txt_TIM_IMPACT_KND 
         Height          =   315
         Left            =   -66210
         MaxLength       =   1
         TabIndex        =   122
         Tag             =   "26"
         Top             =   1140
         Width           =   405
      End
      Begin VB.TextBox txt_TIM_IMPACT_KND_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -65790
         TabIndex        =   175
         Top             =   1140
         Width           =   1695
      End
      Begin VB.TextBox txt_TIM_IMPACT_DIR_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62040
         TabIndex        =   174
         Top             =   1140
         Width           =   1695
      End
      Begin VB.TextBox txt_IMPACT_DIR_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69480
         TabIndex        =   173
         Top             =   1140
         Width           =   1695
      End
      Begin VB.TextBox txt_IMPACT_KND_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73200
         TabIndex        =   172
         Top             =   1140
         Width           =   1695
      End
      Begin VB.TextBox txt_WLD_HARD_NAME 
         Height          =   300
         Left            =   390
         TabIndex        =   171
         Top             =   5880
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txt_HARD_NAME 
         Height          =   300
         Left            =   390
         TabIndex        =   170
         Top             =   6210
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txt_FOAT_RST 
         Height          =   300
         Left            =   6300
         TabIndex        =   169
         Tag             =   "99"
         Top             =   6660
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txt_WLD_BEND_RST 
         Height          =   300
         Left            =   4464
         TabIndex        =   168
         Tag             =   "99"
         Top             =   6675
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txt_BEND_RST 
         Height          =   300
         Left            =   2678
         TabIndex        =   167
         Tag             =   "99"
         Top             =   6660
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txt_IMPACT_KND 
         Height          =   315
         Left            =   -73620
         MaxLength       =   1
         TabIndex        =   120
         Tag             =   "24"
         Top             =   1140
         Width           =   405
      End
      Begin VB.TextBox txt_IMPACT_DIR 
         Height          =   315
         Left            =   -69900
         MaxLength       =   1
         TabIndex        =   121
         Tag             =   "24"
         Top             =   1140
         Width           =   405
      End
      Begin VB.ComboBox cbo_FOAT_RST 
         Height          =   300
         ItemData        =   "AQC0031C.frx":0124
         Left            =   6310
         List            =   "AQC0031C.frx":0131
         TabIndex        =   23
         Tag             =   "19"
         Top             =   7050
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ComboBox cbo_WLD_BEND_RST 
         Height          =   300
         ItemData        =   "AQC0031C.frx":0140
         Left            =   4464
         List            =   "AQC0031C.frx":014D
         TabIndex        =   21
         Tag             =   "17"
         Top             =   7065
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ComboBox cbo_BEND_RST 
         Height          =   300
         ItemData        =   "AQC0031C.frx":015C
         Left            =   2678
         List            =   "AQC0031C.frx":0169
         TabIndex        =   19
         Tag             =   "15"
         Top             =   7050
         Width           =   705
      End
      Begin VB.TextBox txt_RPT_BEND_RST 
         Height          =   300
         Left            =   5402
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "18"
         Top             =   7050
         Width           =   705
      End
      Begin VB.TextBox txt_WLD_HARD_TYP 
         Height          =   300
         Left            =   3556
         TabIndex        =   137
         Tag             =   "16"
         Top             =   6645
         Width           =   705
      End
      Begin VB.TextBox txt_HARD_TYP 
         Height          =   300
         Left            =   1740
         TabIndex        =   0
         Tag             =   "14"
         Top             =   6630
         Width           =   705
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   6
         Left            =   2700
         TabIndex        =   139
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "规    定 总伸长 应    力                                                           MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "规    定 总伸长 应    力                                                           MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2085
         Index           =   7
         Left            =   5400
         TabIndex        =   140
         Top             =   4530
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3678
         _StockProps     =   111
         Caption         =   "反    复弯    曲                                                                              次"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "反    复弯    曲                                                                              次"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   1395
         Index           =   8
         Left            =   4500
         TabIndex        =   141
         Top             =   5205
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   2461
         _StockProps     =   111
         Caption         =   "焊  缝 弯  曲"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "焊  缝 弯  曲"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   1395
         Index           =   9
         Left            =   3585
         TabIndex        =   142
         Top             =   5205
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   2461
         _StockProps     =   111
         Caption         =   "焊  接 硬  度"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "焊  接 硬  度"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2055
         Index           =   10
         Left            =   1770
         TabIndex        =   143
         Top             =   4530
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3625
         _StockProps     =   111
         Caption         =   "硬    度试    验"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "硬    度试    验"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2085
         Index           =   11
         Left            =   6315
         TabIndex        =   144
         Top             =   4530
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3678
         _StockProps     =   111
         Caption         =   "锻    平试    验"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "锻    平试    验"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   13
         Left            =   4500
         TabIndex        =   145
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "规    定 残    余 伸    长 应    力                                                    MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "规    定 残    余 伸    长 应    力                                                    MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   15
         Left            =   6360
         TabIndex        =   146
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "断    后伸长率                  EL                       δ                                  %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "断    后伸长率                  EL                       δ                                  %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   16
         Left            =   5430
         TabIndex        =   147
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "抗    拉 强    度                 TS                      σb                          MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "抗    拉 强    度                 TS                      σb                          MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   17
         Left            =   7245
         TabIndex        =   148
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "断    面 收缩率                 RA                      ψ                               %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "断    面 收缩率                 RA                      ψ                               %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   18
         Left            =   1770
         TabIndex        =   149
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "屈    服 强    度                 YP                      σs                             MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "屈    服 强    度                 YP                      σs                             MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2055
         Index           =   19
         Left            =   2685
         TabIndex        =   150
         Top             =   4530
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3625
         _StockProps     =   111
         Caption         =   "弯    曲试    验"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "弯    曲试    验"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin InDate.ULabel ULabel1 
         Height          =   300
         Index           =   8
         Left            =   3720
         Top             =   4545
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         Caption         =   "焊接试验"
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
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2385
         Index           =   20
         Left            =   11265
         TabIndex        =   151
         Top             =   4530
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   4207
         _StockProps     =   111
         Caption         =   "重    力撕    裂试    验                                                                 %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "重    力撕    裂试    验                                                                 %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2385
         Index           =   21
         Left            =   10215
         TabIndex        =   152
         Top             =   4530
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   4207
         _StockProps     =   111
         Caption         =   "硫化物 腐    蚀 裂    纹                                                                  %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "硫化物 腐    蚀 裂    纹                                                                  %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   1725
         Index           =   22
         Left            =   9120
         TabIndex        =   153
         Top             =   5190
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3043
         _StockProps     =   111
         Caption         =   "  CWR                                                        %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "  CWR                                                        %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   1725
         Index           =   23
         Left            =   7215
         TabIndex        =   154
         Top             =   5190
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3043
         _StockProps     =   111
         Caption         =   "  CSR                                                      %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "  CSR                                                      %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   1725
         Index           =   24
         Left            =   8130
         TabIndex        =   155
         Top             =   5190
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3043
         _StockProps     =   111
         Caption         =   "  CLR                                                         %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "  CLR                                                         %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin InDate.ULabel ULabel1 
         Height          =   300
         Index           =   9
         Left            =   7860
         Top             =   4545
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         Caption         =   "抗氢裂能力"
         Alignment       =   1
         BackColor       =   14804173
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
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   0
         Left            =   14400
         TabIndex        =   156
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "规    定 残    余 伸    长 应    力                                                    MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "规    定 残    余 伸    长 应    力                                                    MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   1
         Left            =   13350
         TabIndex        =   157
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "规    定 非比例 伸    长 应    力                                                      MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "规    定 非比例 伸    长 应    力                                                      MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   2
         Left            =   12285
         TabIndex        =   158
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "断    后伸长率                  EL                       δ                                  %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "断    后伸长率                  EL                       δ                                  %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   3
         Left            =   10170
         TabIndex        =   159
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "抗    拉 强    度                 TS                      σb                          MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "抗    拉 强    度                 TS                      σb                          MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   4
         Left            =   11235
         TabIndex        =   160
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "断    面 收缩率                 RA                      ψ                               %"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "断    面 收缩率                 RA                      ψ                               %"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   5
         Left            =   9120
         TabIndex        =   161
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "屈    服 强    度                 YP                      σs                             MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "屈    服 强    度                 YP                      σs                             MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sidbEdit sdb_HARD_RST 
         Height          =   315
         Left            =   1740
         TabIndex        =   18
         Tag             =   "14"
         Top             =   7050
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_WLD_HARD_RST 
         Height          =   315
         Left            =   3555
         TabIndex        =   20
         Tag             =   "16"
         Top             =   7065
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR_RST 
         Height          =   315
         Left            =   7185
         TabIndex        =   24
         Tag             =   "21"
         Top             =   7050
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR_RST 
         Height          =   315
         Left            =   8100
         TabIndex        =   25
         Tag             =   "21"
         Top             =   7050
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HIC_CWR_RST 
         Height          =   315
         Left            =   9090
         TabIndex        =   26
         Tag             =   "21"
         Top             =   7050
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SSCC_YP_RST 
         Height          =   315
         Left            =   10215
         TabIndex        =   27
         Tag             =   "22"
         Top             =   7050
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_RST1 
         Height          =   315
         Left            =   11205
         TabIndex        =   28
         Tag             =   "23"
         Top             =   7050
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST1 
         Height          =   315
         Left            =   -73410
         TabIndex        =   32
         Tag             =   "24"
         Top             =   2655
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST2 
         Height          =   315
         Left            =   -72600
         TabIndex        =   33
         Tag             =   "24"
         Top             =   2655
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST3 
         Height          =   315
         Left            =   -71790
         TabIndex        =   34
         Tag             =   "24"
         Top             =   2655
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST4 
         Height          =   315
         Left            =   -70980
         TabIndex        =   35
         Tag             =   "24"
         Top             =   2655
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST5 
         Height          =   315
         Left            =   -70170
         TabIndex        =   36
         Tag             =   "24"
         Top             =   2655
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST6 
         Height          =   315
         Left            =   -69360
         TabIndex        =   37
         Tag             =   "24"
         Top             =   2655
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST_AVE 
         Height          =   315
         Left            =   -68550
         TabIndex        =   38
         Tag             =   "24"
         Top             =   2655
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST1 
         Height          =   315
         Left            =   -67380
         TabIndex        =   47
         Tag             =   "26"
         Top             =   2970
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST2 
         Height          =   315
         Left            =   -66570
         TabIndex        =   48
         Tag             =   "26"
         Top             =   2970
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST3 
         Height          =   315
         Left            =   -65760
         TabIndex        =   49
         Tag             =   "26"
         Top             =   2955
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST4 
         Height          =   315
         Left            =   -64950
         TabIndex        =   50
         Tag             =   "26"
         Top             =   2970
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST5 
         Height          =   315
         Left            =   -64140
         TabIndex        =   51
         Tag             =   "26"
         Top             =   2970
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST6 
         Height          =   315
         Left            =   -63330
         TabIndex        =   52
         Tag             =   "26"
         Top             =   2970
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST_AVE 
         Height          =   315
         Left            =   -62520
         TabIndex        =   53
         Tag             =   "26"
         Top             =   2970
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RATE_RST 
         Height          =   315
         Left            =   -61710
         TabIndex        =   54
         Tag             =   "26"
         Top             =   2970
         Width           =   1350
         _Version        =   262145
         _ExtentX        =   2381
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_RST 
         Height          =   315
         Left            =   -71070
         TabIndex        =   79
         Tag             =   "28"
         Top             =   1515
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RMV_CAR_RST 
         Height          =   315
         Left            =   -71070
         TabIndex        =   80
         Tag             =   "29"
         Top             =   2610
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_S_PRINT_RST 
         Height          =   315
         Left            =   -71070
         TabIndex        =   81
         Tag             =   "30"
         Top             =   3240
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_ARST3 
         Height          =   315
         Left            =   -60990
         TabIndex        =   112
         Tag             =   "34"
         Top             =   2370
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_ARST2 
         Height          =   315
         Left            =   -60990
         TabIndex        =   111
         Tag             =   "34"
         Top             =   1965
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_ARST1 
         Height          =   315
         Left            =   -60990
         TabIndex        =   110
         Tag             =   "34"
         Top             =   1500
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BRST2 
         Height          =   315
         Left            =   -60990
         TabIndex        =   115
         Tag             =   "34"
         Top             =   3720
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_ARST4 
         Height          =   315
         Left            =   -60975
         TabIndex        =   113
         Tag             =   "34"
         Top             =   2730
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BRST3 
         Height          =   315
         Left            =   -60990
         TabIndex        =   116
         Tag             =   "34"
         Top             =   4140
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BRST4 
         Height          =   315
         Left            =   -60990
         TabIndex        =   117
         Tag             =   "34"
         Top             =   4560
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BRST1 
         Height          =   315
         Left            =   -60990
         TabIndex        =   114
         Tag             =   "34"
         Top             =   3330
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   1
         Left            =   -74880
         Top             =   630
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   529
         Caption         =   "冲击试验"
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
         Index           =   2
         Left            =   -74850
         Top             =   4050
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   529
         Caption         =   "追加冲击试验"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   13
         Left            =   -71100
         Top             =   1140
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "试样方向"
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
         Index           =   14
         Left            =   -74775
         Top             =   1140
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "类别"
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
         Height          =   330
         Index           =   15
         Left            =   -74775
         Top             =   2655
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         Caption         =   "试验实绩（J）"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   16
         Left            =   -73410
         Top             =   2325
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "1"
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
         Index           =   17
         Left            =   -72600
         Top             =   2325
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "2"
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
         Index           =   18
         Left            =   -71790
         Top             =   2325
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "3"
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
         Index           =   19
         Left            =   -70980
         Top             =   2325
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "4"
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
         Index           =   20
         Left            =   -70170
         Top             =   2325
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "5"
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
         Index           =   21
         Left            =   -69360
         Top             =   2325
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "6"
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
         Index           =   22
         Left            =   -68550
         Top             =   2325
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "平均值"
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
         Index           =   23
         Left            =   -74775
         Top             =   3060
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "断面纤维率(%)"
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   4
         Left            =   -67440
         Top             =   630
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   529
         Caption         =   "时效冲击试验"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   24
         Left            =   -63660
         Top             =   1140
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "试样方向"
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
         Index           =   25
         Left            =   -67380
         Top             =   1140
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "类别"
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
         Index           =   26
         Left            =   -67380
         Top             =   2310
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   556
         Caption         =   "试验实绩（J）"
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
         Index           =   27
         Left            =   -67380
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "1"
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
         Index           =   28
         Left            =   -66570
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "2"
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
         Index           =   29
         Left            =   -65760
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "3"
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
         Index           =   30
         Left            =   -64950
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "4"
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
         Index           =   31
         Left            =   -64140
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "5"
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
         Index           =   32
         Left            =   -63330
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "6"
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
         Index           =   33
         Left            =   -62520
         Top             =   2640
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "平均值"
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
         Height          =   645
         Index           =   34
         Left            =   -61710
         Top             =   2310
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1138
         Caption         =   "断面纤维率(%)"
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
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST1 
         Height          =   315
         Left            =   -73410
         TabIndex        =   56
         Tag             =   "25"
         Top             =   6030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST2 
         Height          =   315
         Left            =   -72600
         TabIndex        =   57
         Tag             =   "25"
         Top             =   6030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST3 
         Height          =   315
         Left            =   -71790
         TabIndex        =   58
         Tag             =   "25"
         Top             =   6030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST4 
         Height          =   315
         Left            =   -70980
         TabIndex        =   59
         Tag             =   "25"
         Top             =   6030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST5 
         Height          =   315
         Left            =   -70170
         TabIndex        =   60
         Tag             =   "25"
         Top             =   6030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST6 
         Height          =   315
         Left            =   -69360
         TabIndex        =   61
         Tag             =   "25"
         Top             =   6030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST_AVE 
         Height          =   315
         Left            =   -68550
         TabIndex        =   62
         Tag             =   "25"
         Top             =   6030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   35
         Left            =   -71100
         Top             =   4590
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "试样方向"
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
         Index           =   36
         Left            =   -74775
         Top             =   4590
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "类别"
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
         Index           =   37
         Left            =   -74775
         Top             =   6030
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "试验实绩（J）"
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
         Index           =   38
         Left            =   -73410
         Top             =   5700
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "1"
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
         Index           =   39
         Left            =   -72600
         Top             =   5700
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "2"
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
         Index           =   40
         Left            =   -71790
         Top             =   5700
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "3"
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
         Index           =   41
         Left            =   -70980
         Top             =   5700
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "4"
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
         Index           =   42
         Left            =   -70170
         Top             =   5700
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "5"
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
         Index           =   43
         Left            =   -69360
         Top             =   5700
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "6"
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
         Index           =   44
         Left            =   -68550
         Top             =   5700
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "平均值"
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
         Index           =   45
         Left            =   -74775
         Top             =   6360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "断面纤维率(%)"
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
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST1 
         Height          =   315
         Left            =   -67350
         TabIndex        =   71
         Tag             =   "27"
         Top             =   6360
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
         Modified        =   -1  'True
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST2 
         Height          =   315
         Left            =   -66540
         TabIndex        =   72
         Tag             =   "27"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST3 
         Height          =   315
         Left            =   -65730
         TabIndex        =   73
         Tag             =   "27"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST4 
         Height          =   315
         Left            =   -64920
         TabIndex        =   74
         Tag             =   "27"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST5 
         Height          =   315
         Left            =   -64110
         TabIndex        =   75
         Tag             =   "27"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST6 
         Height          =   315
         Left            =   -63300
         TabIndex        =   76
         Tag             =   "27"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST_AVE 
         Height          =   315
         Left            =   -62490
         TabIndex        =   77
         Tag             =   "27"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RATE_RST 
         Height          =   315
         Left            =   -61680
         TabIndex        =   78
         Tag             =   "27"
         Top             =   6360
         Width           =   1350
         _Version        =   262145
         _ExtentX        =   2381
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   5
         Left            =   -67440
         Top             =   4050
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   529
         Caption         =   "追加时效冲击试验"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   46
         Left            =   -63660
         Top             =   4590
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "试样方向"
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
         Index           =   47
         Left            =   -67380
         Top             =   4590
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "类别"
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
         Index           =   48
         Left            =   -67350
         Top             =   5700
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   556
         Caption         =   "试验实绩（J）"
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
         Index           =   49
         Left            =   -67350
         Top             =   6030
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "1"
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
         Index           =   50
         Left            =   -66540
         Top             =   6030
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "2"
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
         Index           =   51
         Left            =   -65730
         Top             =   6030
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "3"
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
         Index           =   52
         Left            =   -64920
         Top             =   6030
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "4"
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
         Index           =   53
         Left            =   -64110
         Top             =   6030
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "5"
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
         Index           =   54
         Left            =   -63300
         Top             =   6030
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "6"
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
         Index           =   55
         Left            =   -62490
         Top             =   6030
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "平均值"
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
         Height          =   645
         Index           =   56
         Left            =   -61680
         Top             =   5700
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1138
         Caption         =   "断面纤维率(%)"
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
         Height          =   300
         Index           =   0
         Left            =   -74730
         Top             =   960
         Width           =   1410
         _ExtentX        =   2487
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   57
         Left            =   -74730
         Top             =   1515
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "晶粒度(级)"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   3
         Left            =   -71070
         Top             =   960
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "实绩"
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
         Index           =   34
         Left            =   -73245
         Top             =   960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   529
         Caption         =   "代码"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   58
         Left            =   -74730
         Top             =   2610
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "脱碳层"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Index           =   59
         Left            =   -74730
         Top             =   3210
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "硫印(级)"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Index           =   60
         Left            =   -74730
         Top             =   3840
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "断口检验"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_RST2 
         Height          =   315
         Left            =   -66240
         TabIndex        =   95
         Tag             =   "32"
         Top             =   1995
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_RST3 
         Height          =   315
         Left            =   -66240
         TabIndex        =   97
         Tag             =   "32"
         Top             =   2400
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_RST1 
         Height          =   315
         Left            =   -66240
         TabIndex        =   93
         Tag             =   "32"
         Top             =   1560
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_RST4 
         Height          =   315
         Left            =   -66240
         TabIndex        =   99
         Tag             =   "32"
         Top             =   2790
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_RST5 
         Height          =   315
         Left            =   -66240
         TabIndex        =   101
         Tag             =   "32"
         Top             =   3225
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
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
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Index           =   23
         Left            =   -69900
         Top             =   1560
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "酸浸检验(级)"
         Alignment       =   0
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   6
         Left            =   -68415
         Top             =   960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   529
         Caption         =   "缺陷名称"
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
         Left            =   -69900
         Top             =   960
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   8
         Left            =   -66240
         Top             =   960
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "实绩"
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   0
         Left            =   -69900
         Top             =   3810
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "带状组织(级)"
         Alignment       =   0
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
         Height          =   1770
         Index           =   1
         Left            =   -63780
         Top             =   1410
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   3122
         Caption         =   "粗系"
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
         Height          =   1680
         Index           =   2
         Left            =   -63780
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   2963
         Caption         =   "细系"
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
         Height          =   4230
         Index           =   25
         Left            =   -64620
         Top             =   1410
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   7461
         Caption         =   "非金属"
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
         Index           =   9
         Left            =   -64680
         Top             =   960
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   10
         Left            =   -61020
         Top             =   960
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "实绩"
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
         Index           =   11
         Left            =   -63195
         Top             =   960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   529
         Caption         =   "代码"
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
      Begin CSTextLibCtl.sidbEdit sdb_FRACT_GRD_RST1 
         Height          =   315
         Left            =   -71070
         TabIndex        =   83
         Tag             =   "31"
         Top             =   3840
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_FRACT_GRD_RST2 
         Height          =   315
         Left            =   -71070
         TabIndex        =   85
         Tag             =   "31"
         Top             =   4260
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_FRACT_GRD_RST3 
         Height          =   315
         Left            =   -71070
         TabIndex        =   87
         Tag             =   "31"
         Top             =   4650
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_FRACT_GRD_RST4 
         Height          =   315
         Left            =   -71070
         TabIndex        =   89
         Tag             =   "31"
         Top             =   5040
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_FRACT_GRD_RST5 
         Height          =   315
         Left            =   -71070
         TabIndex        =   91
         Tag             =   "31"
         Top             =   5460
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sivbLB sivbLB1 
         Height          =   2265
         Index           =   29
         Left            =   3600
         TabIndex        =   199
         Top             =   780
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   3995
         _StockProps     =   111
         Caption         =   "规    定 非比例 伸    长 应    力                                                      MPa"
         ForeColor       =   -2147483640
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "规    定 非比例 伸    长 应    力                                                      MPa"
         BorderStyle     =   0
         Alignment       =   1
         BorderEffect    =   2
         ChiselText      =   2
         WordWrap        =   -1  'True
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST1 
         Height          =   315
         Left            =   -73410
         TabIndex        =   39
         Tag             =   "24"
         Top             =   3045
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST2 
         Height          =   315
         Left            =   -72600
         TabIndex        =   40
         Tag             =   "24"
         Top             =   3045
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST3 
         Height          =   315
         Left            =   -71760
         TabIndex        =   41
         Tag             =   "24"
         Top             =   3045
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST4 
         Height          =   315
         Left            =   -70980
         TabIndex        =   42
         Tag             =   "24"
         Top             =   3030
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST5 
         Height          =   315
         Left            =   -70170
         TabIndex        =   43
         Tag             =   "24"
         Top             =   3045
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST6 
         Height          =   315
         Left            =   -69360
         TabIndex        =   44
         Tag             =   "24"
         Top             =   3045
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_AVE_RST 
         Height          =   315
         Left            =   -68550
         TabIndex        =   45
         Tag             =   "24"
         Top             =   3045
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_RST2 
         Height          =   315
         Left            =   11205
         TabIndex        =   29
         Tag             =   "23"
         Top             =   7365
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_RST3 
         Height          =   315
         Left            =   11205
         TabIndex        =   30
         Tag             =   "23"
         Top             =   7680
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_RST_TOP1 
         Height          =   315
         Left            =   -66225
         TabIndex        =   104
         Tag             =   "20"
         Top             =   4725
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_RST_TOP2 
         Height          =   315
         Left            =   -66225
         TabIndex        =   107
         Tag             =   "20"
         Top             =   5085
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_RST_TOP3 
         Height          =   315
         Left            =   -66225
         TabIndex        =   109
         Tag             =   "20"
         Top             =   5460
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_DIST1 
         Height          =   315
         Left            =   -67650
         TabIndex        =   103
         Tag             =   "20"
         Top             =   4725
         Width           =   1320
         _Version        =   262145
         _ExtentX        =   2328
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_DIST2 
         Height          =   315
         Left            =   -67650
         TabIndex        =   106
         Tag             =   "20"
         Top             =   5085
         Width           =   1320
         _Version        =   262145
         _ExtentX        =   2328
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_DIST3 
         Height          =   315
         Left            =   -67650
         TabIndex        =   108
         Tag             =   "20"
         Top             =   5460
         Width           =   1320
         _Version        =   262145
         _ExtentX        =   2328
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   3
         Left            =   -69900
         Top             =   4725
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "淬透性试验"
         Alignment       =   0
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
         Index           =   12
         Left            =   -69900
         Top             =   4350
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   13
         Left            =   -66240
         Top             =   4350
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "实绩"
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
         Index           =   14
         Left            =   -68445
         Top             =   4350
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "类型"
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
         Index           =   15
         Left            =   -67650
         Top             =   4350
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Caption         =   "位置"
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
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST1 
         Height          =   315
         Left            =   -73410
         TabIndex        =   63
         Tag             =   "25"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST2 
         Height          =   315
         Left            =   -72600
         TabIndex        =   64
         Tag             =   "25"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST3 
         Height          =   315
         Left            =   -71790
         TabIndex        =   65
         Tag             =   "25"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST4 
         Height          =   315
         Left            =   -70980
         TabIndex        =   66
         Tag             =   "25"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST5 
         Height          =   315
         Left            =   -70155
         TabIndex        =   67
         Tag             =   "25"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST6 
         Height          =   315
         Left            =   -69360
         TabIndex        =   68
         Tag             =   "25"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_AVE_RST 
         Height          =   315
         Left            =   -68550
         TabIndex        =   69
         Tag             =   "25"
         Top             =   6360
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_YP_RST 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Tag             =   "1"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TS_RST 
         Height          =   315
         Left            =   5400
         TabIndex        =   5
         Tag             =   "2"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RA_RST_1 
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Tag             =   "3"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_EL_RST 
         Height          =   315
         Left            =   6330
         TabIndex        =   6
         Tag             =   "4"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SNPP_EL_RST 
         Height          =   315
         Left            =   3570
         TabIndex        =   3
         Tag             =   "5"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SP_EL_RST 
         Height          =   315
         Left            =   4470
         TabIndex        =   4
         Tag             =   "6"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SG_EL_RST 
         Height          =   315
         Left            =   2655
         TabIndex        =   2
         Tag             =   "7"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_RST 
         Height          =   315
         Left            =   9090
         TabIndex        =   12
         Tag             =   "8"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_RST 
         Height          =   315
         Left            =   10140
         TabIndex        =   13
         Tag             =   "9"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST 
         Height          =   315
         Left            =   11190
         TabIndex        =   14
         Tag             =   "10"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_RST 
         Height          =   315
         Left            =   12240
         TabIndex        =   15
         Tag             =   "111"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_RST 
         Height          =   315
         Left            =   13350
         TabIndex        =   16
         Tag             =   "12"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
         Height          =   315
         Left            =   14400
         TabIndex        =   17
         Tag             =   "13"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_YR_RST 
         Height          =   315
         Left            =   8100
         TabIndex        =   11
         Tag             =   "35"
         Top             =   3120
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RA_RST_2 
         Height          =   315
         Left            =   7200
         TabIndex        =   8
         Tag             =   "3"
         Top             =   3450
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RA_RST_3 
         Height          =   315
         Left            =   7200
         TabIndex        =   9
         Tag             =   "3"
         Top             =   3780
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RA_RST_AVE 
         Height          =   315
         Left            =   7200
         TabIndex        =   10
         Tag             =   "3"
         Top             =   4110
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   62
         Left            =   -74775
         Top             =   1770
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "试片尺寸"
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
         Index           =   63
         Left            =   -67380
         Top             =   1770
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "试片尺寸"
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
         Index           =   64
         Left            =   -74775
         Top             =   5160
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "试片尺寸"
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
         Index           =   65
         Left            =   -67380
         Top             =   5160
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "试片尺寸"
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
      Begin CSTextLibCtl.sidbEdit txt_BELT_STR_GRD_RST 
         Height          =   300
         Left            =   -66240
         TabIndex        =   212
         Tag             =   "33"
         Top             =   3810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Text            =   " 0.0"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   13
         CurNumDataChars =   0
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_OST_GRAIN_SIZE_RST 
         Height          =   315
         Left            =   -71070
         TabIndex        =   216
         Tag             =   "36"
         Top             =   1980
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   66
         Left            =   -74730
         Top             =   1980
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "奥氏体晶粒度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   330
         Index           =   0
         Left            =   -63780
         Top             =   4950
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         Caption         =   "DS"
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
      Begin CSTextLibCtl.sidbEdit sdb_DS_GRD 
         Height          =   315
         Left            =   -60990
         TabIndex        =   217
         Tag             =   "34"
         Top             =   4995
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel87 
         Height          =   330
         Index           =   3
         Left            =   -63780
         Top             =   5310
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         Caption         =   "TIN"
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
      Begin CSTextLibCtl.sidbEdit sdb_TIN_GRD 
         Height          =   315
         Left            =   -60990
         TabIndex        =   218
         Tag             =   "34"
         Top             =   5355
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "试验实绩"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   201
         Top             =   3180
         Width           =   1440
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   -69930
         X2              =   -65265
         Y1              =   5895
         Y2              =   5895
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   -74760
         X2              =   -70110
         Y1              =   5895
         Y2              =   5895
      End
      Begin VB.Line Line5 
         Index           =   15
         X1              =   4365
         X2              =   4365
         Y1              =   750
         Y2              =   4470
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   -64710
         X2              =   -60015
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   -64710
         X2              =   -60060
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   -69930
         X2              =   -65280
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   -69930
         X2              =   -65280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   -74760
         X2              =   -70110
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   -74760
         X2              =   -70110
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -64725
         X2              =   -60000
         Y1              =   5700
         Y2              =   5700
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   -63240
         X2              =   -60015
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   -60015
         X2              =   -60015
         Y1              =   840
         Y2              =   5700
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   -61050
         X2              =   -61050
         Y1              =   960
         Y2              =   5700
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   -63240
         X2              =   -63240
         Y1              =   960
         Y2              =   5700
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   -64725
         X2              =   -64725
         Y1              =   840
         Y2              =   5700
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   -65280
         X2              =   -65280
         Y1              =   840
         Y2              =   5895
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   -69915
         X2              =   -65250
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   -66285
         X2              =   -66285
         Y1              =   840
         Y2              =   4260
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   -69930
         X2              =   -65280
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   -68460
         X2              =   -68460
         Y1              =   840
         Y2              =   5880
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   -69945
         X2              =   -69945
         Y1              =   855
         Y2              =   5910
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   -70110
         X2              =   -70110
         Y1              =   840
         Y2              =   5900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   -74760
         X2              =   -74760
         Y1              =   840
         Y2              =   5900
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   -74760
         X2              =   -70110
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   -74760
         X2              =   -70110
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -74760
         X2              =   -70110
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -71115
         X2              =   -71115
         Y1              =   840
         Y2              =   5900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   24
         X1              =   -73290
         X2              =   -73290
         Y1              =   840
         Y2              =   5900
      End
      Begin VB.Shape Shape3 
         Height          =   2500
         Left            =   -67440
         Top             =   4380
         Width           =   7185
      End
      Begin VB.Shape Shape1 
         Height          =   2500
         Index           =   1
         Left            =   -74850
         Top             =   4380
         Width           =   7185
      End
      Begin VB.Shape Shape2 
         Height          =   2500
         Left            =   -67440
         Top             =   960
         Width           =   7185
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "高温拉伸试验"
         Height          =   180
         Index           =   5
         Left            =   9030
         TabIndex        =   166
         Top             =   510
         Width           =   6030
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "试验项目"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   165
         Top             =   4530
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "试验实绩"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   164
         Top             =   7050
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "拉伸试验"
         Height          =   180
         Index           =   1
         Left            =   1740
         TabIndex        =   163
         Top             =   510
         Width           =   7110
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "试验项目"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   162
         Top             =   540
         Width           =   1440
      End
      Begin VB.Line Line5 
         Index           =   14
         X1              =   -75000
         X2              =   -75000
         Y1              =   870
         Y2              =   6975
      End
      Begin VB.Line Line5 
         Index           =   13
         X1              =   7980
         X2              =   7980
         Y1              =   5025
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   12
         X1              =   8910
         X2              =   8910
         Y1              =   5025
         Y2              =   8045
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   7110
         X2              =   9975
         Y1              =   5010
         Y2              =   5010
      End
      Begin VB.Line Line5 
         Index           =   11
         X1              =   4350
         X2              =   4350
         Y1              =   5025
         Y2              =   8045
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   3465
         X2              =   5295
         Y1              =   5025
         Y2              =   5025
      End
      Begin VB.Line Line5 
         Index           =   10
         X1              =   13182
         X2              =   13182
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   9
         X1              =   12120
         X2              =   12120
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   8
         X1              =   7995
         X2              =   7995
         Y1              =   750
         Y2              =   4470
      End
      Begin VB.Line Line5 
         Index           =   7
         X1              =   14250
         X2              =   14250
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   6
         X1              =   11040
         X2              =   11040
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   5
         X1              =   9978
         X2              =   9978
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   4
         X1              =   7092
         X2              =   7092
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   3
         X1              =   6180
         X2              =   6180
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   5278
         X2              =   5278
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   3464
         X2              =   3464
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   2557
         X2              =   2557
         Y1              =   750
         Y2              =   8045
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   0
         X2              =   15210
         Y1              =   6990
         Y2              =   6990
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   0
         X2              =   15210
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   30
         X2              =   15240
         Y1              =   4470
         Y2              =   4470
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   1650
         X2              =   15240
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   1
         X1              =   8910
         X2              =   8895
         Y1              =   450
         Y2              =   4470
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   1650
         X2              =   1650
         Y1              =   450
         Y2              =   8045
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   3
         X1              =   -75000
         X2              =   -59760
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   15240
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Shape Shape1 
         Height          =   2505
         Index           =   0
         Left            =   -74850
         Top             =   960
         Width           =   7185
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "试样编号"
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
      Left            =   4410
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "取样位置"
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
      Index           =   2
      Left            =   210
      Top             =   555
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   556
      Caption         =   "钢种"
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
      Index           =   3
      Left            =   2160
      Top             =   555
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      Caption         =   "炉号"
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
      Index           =   4
      Left            =   4890
      Top             =   555
      Width           =   2010
      _ExtentX        =   3545
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   5
      Left            =   8160
      Top             =   555
      Width           =   1950
      _ExtentX        =   3440
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   6
      Left            =   10110
      Top             =   555
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "序列号"
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
      Index           =   7
      Left            =   11430
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "订单用途"
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
      Index           =   10
      Left            =   12660
      Top             =   555
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "订单厚度"
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
      Index           =   11
      Left            =   13860
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "订单宽度"
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
      Index           =   12
      Left            =   3600
      Top             =   555
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "生产日期"
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
   Begin InDate.ULabel lbl_STLGRD 
      Height          =   345
      Left            =   210
      Top             =   855
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_HEAT_NO 
      Height          =   345
      Left            =   2160
      Top             =   855
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_STDSPEC 
      Height          =   345
      Left            =   4890
      Top             =   855
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_ORD_NO 
      Height          =   345
      Left            =   8160
      Top             =   855
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_ORD_ITEM 
      Height          =   345
      Left            =   10110
      Top             =   855
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_ENDUSE_CD 
      Height          =   345
      Left            =   11430
      Top             =   855
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_ORD_THK 
      Height          =   345
      Left            =   12660
      Top             =   855
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   ""
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_ORD_WID 
      Height          =   345
      Left            =   13860
      Top             =   855
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel lbl_Cut_DD 
      Height          =   345
      Left            =   3600
      Top             =   855
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   61
      Left            =   6900
      Top             =   555
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "发布年度"
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
   Begin InDate.ULabel lbl_STD_YY 
      Height          =   345
      Left            =   6900
      Top             =   855
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin VB.TextBox txt_SMP_NO_P 
      Height          =   285
      Left            =   6210
      TabIndex        =   206
      Top             =   780
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox txt_SMP_CUT_LOC_P 
      Height          =   270
      Left            =   5670
      TabIndex        =   207
      Top             =   840
      Visible         =   0   'False
      Width           =   765
   End
End
Attribute VB_Name = "AQC0031C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      材质试验实绩输入
'-- Program ID        AQC0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.8. 18
'-- Description       材质试验实绩输入
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection


Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection




Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'TOP
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
              Call Gp_Ms_Collection(txt_smp_no_p, "p", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(lbl_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(lbl_Cut_DD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(lbl_HEAT_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(lbl_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(lbl_STD_YY, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(lbl_ORD_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(lbl_ORD_ITEM, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
             Call Gp_Ms_Collection(lbl_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(lbl_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(lbl_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

    'MASTER2 Collection
     Mc2.Add Item:="AQC0031C.P_REFER_HEAD", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"

'----------------------------------------------------------------------------------------------------------------------------------------------------------------

                Call Gp_Ms_Collection(txt_smp_no_p, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_SMP_CUT_LOC_P, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'拉伸／高温拉伸／其它 - TAB 1
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                
                Call Gp_Ms_Collection(sdb_YP_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_SG_EL_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_SNPP_EL_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_SP_EL_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_TS_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_EL_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_RA_RST_1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_RA_RST_2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'10
              Call Gp_Ms_Collection(sdb_RA_RST_3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_RA_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_YR_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HGT_YP_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HGT_TS_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HGT_RA_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HGT_EL_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HARD_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20
             Call Gp_Ms_Collection(txt_HARD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_HARD_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_BEND_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_WLD_HARD_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_WLD_HARD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_WLD_HARD_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_WLD_BEND_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_RPT_BEND_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_FOAT_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_HIC_CSR_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'30
           Call Gp_Ms_Collection(sdb_HIC_CLR_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_HIC_CWR_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_SSCC_YP_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_DWTT_YP_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_DWTT_YP_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_DWTT_YP_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'冲击／时效冲击 - TAB 2
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
     
            Call Gp_Ms_Collection(txt_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'40
           Call Gp_Ms_Collection(Cob_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'50
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_IMPACT_RATE_AVE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
          Call Gp_Ms_Collection(txt_A_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_A_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_A_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_A_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'60
         Call Gp_Ms_Collection(Cob_A_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_A_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'70
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_AVE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
        Call Gp_Ms_Collection(txt_TIM_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_TIM_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_TIM_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_TIM_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Cob_TIM_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'80
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_TIM_IMPACT_RATE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

      Call Gp_Ms_Collection(txt_A_TIM_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_A_TIM_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'90
      Call Gp_Ms_Collection(txt_A_TIM_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_A_TIM_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Cob_A_TIM_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'100
 Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RATE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'金相检验 - TAB 3
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call Gp_Ms_Collection(sdb_GRAIN_SIZE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_RMV_CAR_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_RMV_CAR_TYP_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_RMV_CAR_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_S_PRINT_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_FRACT_GRD_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_FRACT_GRD_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_FRACT_GRD_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_FRACT_GRD_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD5_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_FRACT_GRD_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
                    
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_ACD_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_ACD_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_ACD_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_ACD_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP5_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_ACD_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
       Call Gp_Ms_Collection(txt_BELT_STR_GRD_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
            Call Gp_Ms_Collection(txt_JOMINY_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_JOMINY_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_JOMINY_DIST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_JOMINY_DIST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_JOMINY_DIST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_JOMINY_RST_TOP1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_JOMINY_RST_TOP2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_JOMINY_RST_TOP3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
                
         Call Gp_Ms_Collection(txt_NON_METAL_ACD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_ARST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_ACD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_ARST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_ACD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_ARST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_ACD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_ARST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BRST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BRST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BRST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BRST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_SAVE_CASE, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
'----------------------------------------------------------- Master End ------------------------------------------------------------------------------------
                    Call Gp_Ms_Collection(txt_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_INS_EMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'  SUN BIN ADD 20090514
     Call Gp_Ms_Collection(sdb_OST_GRAIN_SIZE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_DS_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_TIN_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
     Mc1.Add Item:="AQC0031C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQC0031C.P_REFER", Key:="P-R"
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


Private Sub cbo_BEND_RST_Change()
    If cbo_BEND_RST.Visible = False Then Exit Sub
    Select Case Trim(cbo_BEND_RST.Text)
        Case "Yes"
            txt_BEND_RST.Text = "Y"
        Case "No"
            txt_BEND_RST.Text = "N"
        Case Else
            txt_BEND_RST.Text = ""
    End Select
End Sub

Private Sub cbo_FOAT_RST_Change()
        If cbo_FOAT_RST.Visible = False Then Exit Sub
        Select Case Trim(cbo_FOAT_RST.Text)
        Case "Yes"
            txt_FOAT_RST.Text = "Y"
        Case "No"
            txt_FOAT_RST.Text = "N"
        Case Else
            txt_FOAT_RST.Text = ""
    End Select

End Sub

Private Sub cbo_WLD_BEND_RST_Change()
        If cbo_WLD_BEND_RST.Visible = False Then Exit Sub
        Select Case Trim(cbo_WLD_BEND_RST.Text)
        Case "Yes"
            txt_WLD_BEND_RST.Text = "Y"
        Case "No"
            txt_WLD_BEND_RST.Text = "N"
        Case Else
            txt_WLD_BEND_RST.Text = ""
    End Select

End Sub

Private Sub Cob_A_IMPACT_SIZE_Change()
    Call Impact_Size_Cob_Select(Cob_A_IMPACT_SIZE, TXT_A_IMPACT_SIZE_CD)
End Sub

Private Sub Cob_A_TIM_IMPACT_SIZE_Change()

    Call Impact_Size_Cob_Select(Cob_A_TIM_IMPACT_SIZE, TXT_A_TIM_IMPACT_SIZE_CD)

End Sub

Private Sub Cob_IMPACT_SIZE_Change()
        
    Call Impact_Size_Cob_Select(Cob_IMPACT_SIZE, TXT_IMPACT_SIZE_CD)
       
End Sub

Private Sub Cob_TIM_IMPACT_SIZE_Change()

    Call Impact_Size_Cob_Select(Cob_TIM_IMPACT_SIZE, TXT_TIM_IMPACT_SIZE_CD)

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
'        Case "txt_SMP_CUT_LOC"      '取样位置
'            sCode = "Q0042"
        
        Case "txt_HARD_TYP"         '硬度试验
            sCode = "Q0010"
            
        Case "txt_WLD_HARD_TYP"     '焊接硬度
            sCode = "Q0011"
        
        Case "txt_JOMINY_TYP"      '淬透性试验
            sCode = "Q0012"
            
        Case "txt_IMPACT_KND"      '冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_IMPACT_KND_NAME
            
        Case "txt_IMPACT_DIR"      '冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_IMPACT_DIR_NAME
            
        Case "txt_A_IMPACT_KND"      '追加冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_A_IMPACT_KND_NAME
            
        Case "txt_A_IMPACT_DIR"      '追加冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_A_IMPACT_DIR_NAME
            
        Case "txt_TIM_IMPACT_KND"      '时效冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_TIM_IMPACT_KND_NAME
            
        Case "txt_TIM_IMPACT_DIR"      '时效冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_TIM_IMPACT_DIR_NAME
            
        Case "txt_A_TIM_IMPACT_KND"      '追加时效冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_A_TIM_IMPACT_KND_NAME
            
        Case "txt_A_TIM_IMPACT_DIR"      '追加时效冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_A_TIM_IMPACT_DIR_NAME
            
            
        Case "txt_RMV_CAR_TYP"          '脱碳层
            sCode = "Q0015"
            Set oCodeName = txt_RMV_CAR_TYP_NAME
            
        Case "txt_FRACT_NAME_CD1"       '断口检验 - 1
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD1_NAME
            
        Case "txt_FRACT_NAME_CD2"       '断口检验 - 2
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD2_NAME
            
        Case "txt_FRACT_NAME_CD3"       '断口检验 - 3
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD3_NAME
            
        Case "txt_FRACT_NAME_CD4"       '断口检验 - 4
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD4_NAME
            
        Case "txt_FRACT_NAME_CD5"       '断口检验 - 5
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD5_NAME
            
        
        Case "txt_ACD_DFT_TYP1"         '酸浸检验(级) - 1
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP1_NAME
            
        Case "txt_ACD_DFT_TYP2"         '酸浸检验(级) - 2
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP2_NAME
            
        Case "txt_ACD_DFT_TYP3"         '酸浸检验(级) - 3
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP3_NAME
            
        Case "txt_ACD_DFT_TYP4"         '酸浸检验(级) - 4
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP4_NAME
                        
        Case "txt_ACD_DFT_TYP5"         '酸浸检验(级) - 5
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP5_NAME
            
            
        Case "txt_NON_METAL_ACD1"         '非金属夹杂 - 粗系 - 1
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD1_NAME
            
        Case "txt_NON_METAL_ACD2"         '非金属夹杂 - 粗系 - 2
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD2_NAME
            
        Case "txt_NON_METAL_ACD3"         '非金属夹杂 - 粗系 - 3
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD3_NAME
            
        Case "txt_NON_METAL_ACD4"         '非金属夹杂 - 粗系 - 4
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD4_NAME
            
        Case "txt_NON_METAL_BCD1"         '非金属夹杂 - 细系 - 1
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD1_NAME
            
        Case "txt_NON_METAL_BCD2"         '非金属夹杂 - 细系 - 2
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD2_NAME
            
        Case "txt_NON_METAL_BCD3"         '非金属夹杂 - 细系 - 3
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD3_NAME
            
        Case "txt_NON_METAL_BCD4"         '非金属夹杂 - 细系 - 4
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD4_NAME
        
        Case Else
            Exit Sub
            
    End Select
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


Private Sub cbo_BEND_RST_Click()
    Select Case Trim(cbo_BEND_RST.Text)
        Case "Yes"
            txt_BEND_RST.Text = "Y"
        Case "No"
            txt_BEND_RST.Text = "N"
        Case Else
            txt_BEND_RST.Text = ""
    End Select
End Sub

Private Sub cbo_WLD_BEND_RST_Click()
        Select Case Trim(cbo_WLD_BEND_RST.Text)
        Case "Yes"
            txt_WLD_BEND_RST.Text = "Y"
        Case "No"
            txt_WLD_BEND_RST.Text = "N"
        Case Else
            txt_WLD_BEND_RST.Text = ""
    End Select
End Sub

Private Sub cbo_FOAT_RST_Click()
        Select Case Trim(cbo_FOAT_RST.Text)
        Case "Yes"
            txt_FOAT_RST.Text = "Y"
        Case "No"
            txt_FOAT_RST.Text = "N"
        Case Else
            txt_FOAT_RST.Text = ""
    End Select
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = 19 Or KeyAscii = 10 Then
        KeyAscii = 0
        Call Form_Pro
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

    Screen.MousePointer = vbDefault
    Op_ONLY.Value = True
    txt_SAVE_CASE.Text = "O"
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

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
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    lbl_STLGRD.Caption = ""
    lbl_HEAT_NO.Caption = ""
    lbl_Cut_DD.Caption = ""
    lbl_STDSPEC.Caption = ""
    lbl_STD_YY.Caption = ""
    lbl_ORD_NO.Caption = ""
    lbl_ORD_ITEM.Caption = ""
    lbl_ENDUSE_CD.Caption = ""
    lbl_ORD_WID.Caption = ""
    lbl_ORD_THK.Caption = ""
    'Op_CHAGE.Value = True
    Op_ONLY.Value = True
    txt_SAVE_CASE.Text = "O"
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()
        
    Dim sMesg As String
    
        Call Form_Cls
    
        If sbtn_SMP_TYPE_SELECT.Value = True Then
            txt_SMP_NO.Text = Mid(txt_SMP_NO, 1, 12) + "00"
        Else
            If Mid(txt_SMP_NO.Text, 13, 2) = "00" Then
                Call MsgBox("现在只能输入非作普样，如果要输入作普样请点击-作普样录入按钮！", vbOKOnly, "系统提示")
                Exit Sub
            End If
        End If
    
        If Gf_Ms_Refer(M_CN1, Mc2, Mc1("nControl"), Mc1("mControl")) Then
            Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
           
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            
        End If
            
    Call subComboValueCheck
    Call subItemLock(txt_SMP_NO.Text, lbl_STDSPEC.Caption, lbl_STD_YY.Caption)
    Call subCODENAMElOCK
    
    Select Case SSTab1.Tab
    Case 0
        If sdb_YP_RST.Visible = True Then sdb_YP_RST.SetFocus
    Case 1
        If TXT_IMPACT_SIZE_CD.Visible = True Then TXT_IMPACT_SIZE_CD.SetFocus
    Case 2
        If sdb_GRAIN_SIZE_RST.Visible = True Then sdb_GRAIN_SIZE_RST.SetFocus
    End Select
    
    
    If Val(lbl_ORD_THK.Caption) >= 12 Then
       If TXT_IMPACT_SIZE_CD.Text = "" Then
            TXT_IMPACT_SIZE_CD.Text = "3"
       End If
    End If
    
End Sub

Public Sub Form_Pro()
       
    If Gf_Mc_Authority(sAuthority, Mc1) Then
          
        txt_INS_EMP.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub



'---------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------- Common Code ( F4 Popup ) -------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------

Private Sub subComboValueCheck()
    
    Select Case txt_BEND_RST.Text
        Case "Y"
            cbo_BEND_RST.ListIndex = 1
        Case "N"
            cbo_BEND_RST.ListIndex = 2
        Case Else
            cbo_BEND_RST.ListIndex = 1
    End Select
    
    
    Select Case txt_WLD_BEND_RST.Text
        Case "Y"
            cbo_WLD_BEND_RST.ListIndex = 1
        Case "N"
            cbo_WLD_BEND_RST.ListIndex = 2
        Case Else
            cbo_WLD_BEND_RST.ListIndex = 1
    End Select
    
    Select Case txt_FOAT_RST.Text
        Case "Y"
            cbo_FOAT_RST.ListIndex = 1
        Case "N"
            cbo_FOAT_RST.ListIndex = 2
        Case Else
            cbo_FOAT_RST.ListIndex = 1
    End Select
    
End Sub

Private Sub subItemLock(ByVal sSMP_NO As String, ByVal sSPEC_NO As String, ByVal sYY As String)
    Dim sQuery, sSQL As String
    Dim arrayRecord  As Variant
    Dim AdoRs        As adodb.Recordset
    Dim icount       As Integer
    Dim iarrCOUNT    As Integer
    Dim sKnd         As String
 
 On Error GoTo Error_Rtn
    
    If sSPEC_NO = "" Or Len(Trim(sSPEC_NO)) = 0 Then
        Call subControlLock(arrayRecord, True, Mc1("iControl"))
        Exit Sub
    End If
    
    sSQL = "         SELECT *  "
    sSQL = sSQL + "  From   QP_QLTY_MATR A, QP_TEST_HEAD B"
    sSQL = sSQL + "  Where  B.SMP_NO   =" + "'" + Trim(txt_SMP_NO) + "'"
    sSQL = sSQL + "  AND    A.ORD_NO   = B.ORD_NO"
    sSQL = sSQL + "  AND    A.ORD_ITEM = B.ORD_ITEM"
    sSQL = sSQL + "  AND    KND        = '2'"
    
    Set AdoRs = New adodb.Recordset
    AdoRs.Open sSQL, M_CN1, adOpenKeyset
    
    If AdoRs.EOF And AdoRs.BOF Then
       sKnd = "1"
    Else
       sKnd = "2"
    End If
    
    AdoRs.Close
    
    sQuery = "{call AQC0031C.P_MART_ITEM_SELECT('" + sSMP_NO + "','" + sSPEC_NO + "','" + sYY + "')}"

    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not (AdoRs.BOF And AdoRs.EOF) Then
        arrayRecord = AdoRs.GetRows
    Else
        GoTo Error_Rtn
    End If
    
    AdoRs.Close
    
    If sKnd = "1" Then
            
        Call subControlLock(arrayRecord, False, Mc1("iControl"))
    
    Else
        For iarrCOUNT = 0 To UBound(arrayRecord, 1)
            If Val(arrayRecord(iarrCOUNT, 0)) = Val(arrayRecord(iarrCOUNT, 1)) Then
                arrayRecord(iarrCOUNT, 0) = arrayRecord(iarrCOUNT, 0)
            Else
                If Val(arrayRecord(iarrCOUNT, 1)) > 0 Then
                    arrayRecord(iarrCOUNT, 0) = arrayRecord(iarrCOUNT, 1)
                Else
                    arrayRecord(iarrCOUNT, 0) = arrayRecord(iarrCOUNT, 0)
                End If
            End If
        Next
        
        Call subControlLock(arrayRecord, False, Mc1("iControl"))
        
    End If
    
    Set AdoRs = Nothing
    Set arrayRecord = Nothing

Error_Rtn:
    
    Set AdoRs = Nothing
    Set arrayRecord = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub subControlLock(ByVal vARRAY As Variant, ByVal bAllLock As Boolean, ByVal iCtrl As Collection)
    Dim icount       As Integer
    Dim iarrCOUNT    As Integer
    
    If bAllLock Then
        For icount = 1 To iCtrl.COUNT
            iCtrl.Item(icount).Visible = False
        Next
    Else
        
        For icount = 1 To iCtrl.COUNT
            If iCtrl.Item(icount).Tag <> 99 And iCtrl.Item(icount).Tag <> "INS_EMP" Then
                    
                    For iarrCOUNT = 0 To UBound(vARRAY, 1)
                        
                        If Val(iCtrl.Item(icount).Tag) = Val(vARRAY(iarrCOUNT, 0)) Then
                            iCtrl.Item(icount).Visible = True
                            Exit For
                        Else
                            iCtrl.Item(icount).Visible = False
                        End If
                        
                    Next
            
            End If
        Next
        If cbo_BEND_RST.Visible = False Then
           txt_BEND_RST.Text = ""
        Else
           If txt_BEND_RST.Text = "" Then
              txt_BEND_RST.Text = "Y"
           End If
        End If
        If cbo_WLD_BEND_RST.Visible = False Then
           txt_WLD_BEND_RST.Text = ""
        Else
           If txt_WLD_BEND_RST.Text = "" Then
              txt_WLD_BEND_RST.Text = "Y"
           End If
        End If
        If cbo_FOAT_RST.Visible = False Then
           txt_FOAT_RST.Text = ""
        Else
           If txt_FOAT_RST.Text = "" Then
              txt_FOAT_RST.Text = "Y"
           End If
        End If
    
    End If


End Sub


Private Sub subCODENAMElOCK()

'   cbo_BEND_RST.Visible = txt_BEND_RST.Visible
'   cbo_WLD_BEND_RST.Visible = txt_WLD_BEND_RST.Visible
'   cbo_FOAT_RST.Visible = txt_FOAT_RST.Visible
   txt_IMPACT_KND_NAME.Visible = txt_IMPACT_KND.Visible
   txt_IMPACT_DIR_NAME.Visible = txt_IMPACT_DIR.Visible
   TXT_IMPACT_SIZE_CD.Visible = Cob_IMPACT_SIZE.Visible
   txt_TIM_IMPACT_KND_NAME.Visible = txt_TIM_IMPACT_KND.Visible
   txt_TIM_IMPACT_DIR_NAME.Visible = txt_TIM_IMPACT_DIR.Visible
   TXT_TIM_IMPACT_SIZE_CD.Visible = Cob_TIM_IMPACT_SIZE.Visible
   txt_A_IMPACT_KND_NAME.Visible = txt_A_IMPACT_KND.Visible
   txt_A_IMPACT_DIR_NAME.Visible = txt_A_IMPACT_DIR.Visible
   TXT_A_IMPACT_SIZE_CD.Visible = Cob_A_IMPACT_SIZE.Visible
   txt_A_TIM_IMPACT_KND_NAME.Visible = txt_A_TIM_IMPACT_KND.Visible
   txt_A_TIM_IMPACT_DIR_NAME.Visible = txt_A_TIM_IMPACT_DIR.Visible
   TXT_A_TIM_IMPACT_SIZE_CD.Visible = Cob_A_TIM_IMPACT_SIZE.Visible
   txt_RMV_CAR_TYP_NAME.Visible = txt_RMV_CAR_TYP.Visible
   txt_FRACT_NAME_CD1_NAME.Visible = txt_FRACT_NAME_CD1.Visible
   txt_FRACT_NAME_CD2_NAME.Visible = txt_FRACT_NAME_CD2.Visible
   txt_FRACT_NAME_CD3_NAME.Visible = txt_FRACT_NAME_CD3.Visible
   txt_FRACT_NAME_CD4_NAME.Visible = txt_FRACT_NAME_CD4.Visible
   txt_FRACT_NAME_CD5_NAME.Visible = txt_FRACT_NAME_CD5.Visible
   txt_ACD_DFT_TYP1_NAME.Visible = txt_ACD_DFT_TYP1.Visible
   txt_ACD_DFT_TYP2_NAME.Visible = txt_ACD_DFT_TYP2.Visible
   txt_ACD_DFT_TYP3_NAME.Visible = txt_ACD_DFT_TYP3.Visible
   txt_ACD_DFT_TYP4_NAME.Visible = txt_ACD_DFT_TYP4.Visible
   txt_ACD_DFT_TYP5_NAME.Visible = txt_ACD_DFT_TYP5.Visible
   txt_NON_METAL_ACD1_NAME.Visible = txt_NON_METAL_ACD1.Visible
   txt_NON_METAL_ACD2_NAME.Visible = txt_NON_METAL_ACD2.Visible
   txt_NON_METAL_ACD3_NAME.Visible = txt_NON_METAL_ACD3.Visible
   txt_NON_METAL_ACD4_NAME.Visible = txt_NON_METAL_ACD4.Visible
   txt_NON_METAL_BCD1_NAME.Visible = txt_NON_METAL_BCD1.Visible
   txt_NON_METAL_BCD2_NAME.Visible = txt_NON_METAL_BCD2.Visible
   txt_NON_METAL_BCD3_NAME.Visible = txt_NON_METAL_BCD3.Visible
   txt_NON_METAL_BCD4_NAME.Visible = txt_NON_METAL_BCD4.Visible



End Sub

Private Sub Op_CHAGE_Click()
    If Trim(txt_SAVE_CASE.Text) <> "1" Then
        txt_SAVE_CASE.Text = 1
    End If
End Sub

Private Sub Option2_Click()
    
End Sub

Private Sub Op_ONLY_Click()
    If Trim(txt_SAVE_CASE.Text) <> "0" Then
        txt_SAVE_CASE.Text = 0
    End If
End Sub

Private Sub sbtn_SMP_TYPE_SELECT_Click(Value As Integer)
    If Value = True Then
        sbtn_SMP_TYPE_SELECT.Caption = "录入状态：作普样录入"
        sbtn_SMP_TYPE_SELECT.ForeColor = &HFF0000
    Else
        sbtn_SMP_TYPE_SELECT.Caption = "录入状态： 常规样录入"
        sbtn_SMP_TYPE_SELECT.ForeColor = &HFF
    End If
    txt_SMP_NO.SetFocus
End Sub

Private Sub sdb_A_IMPACT_RATE_RST3_lostfocus()
         sdb_A_IMPACT_RATE_AVE_RST.SetFocus
End Sub

Private Sub sdb_A_IMPACT_RST3_lostfocus()
    sdb_A_IMPACT_RST_AVE.SetFocus
End Sub

Private Sub sdb_A_TIM_IMPACT_RST3_lostfocus()
    sdb_A_TIM_IMPACT_RST_AVE.SetFocus
End Sub

Private Sub sdb_IMPACT_RATE_RST3_lostfocus()
    If sdb_IMPACT_RATE_RST3.Value Then
        sdb_IMPACT_RATE_AVE_RST.SetFocus
    End If
End Sub

Private Sub sdb_IMPACT_RST3_LostFocus()
   If sdb_IMPACT_RST_AVE.Visible = True Then
      sdb_IMPACT_RST_AVE.SetFocus
    End If
End Sub

Private Sub sdb_TIM_IMPACT_RST3_LostFocus()
    If sdb_TIM_IMPACT_RST_AVE.Visible = True Then
       sdb_TIM_IMPACT_RST_AVE.SetFocus
    End If
End Sub

Private Sub TXT_A_IMPACT_SIZE_CD_Change()
   Call Impact_Size_Text_Select(Cob_A_IMPACT_SIZE, TXT_A_IMPACT_SIZE_CD)
End Sub

Private Sub TXT_A_TIM_IMPACT_SIZE_CD_Change()

    Call Impact_Size_Text_Select(Cob_A_TIM_IMPACT_SIZE, TXT_A_TIM_IMPACT_SIZE_CD)

End Sub

Private Sub TXT_IMPACT_SIZE_CD_Change()
    Call Impact_Size_Text_Select(Cob_IMPACT_SIZE, TXT_IMPACT_SIZE_CD)
End Sub

Private Sub txt_SMP_CUT_LOC_Change()
    txt_SMP_CUT_LOC_P.Text = txt_SMP_CUT_LOC.Text
End Sub

Private Sub txt_SMP_CUT_LOC_LostFocus()
    Call Form_Ref
End Sub

Private Sub Impact_Size_Cob_Select(oCob As ComboBox, oEdit As TextBox)
     
     Select Case Trim(oCob.Text)
            Case "5*10*55"
                oEdit.Text = "1"
            Case "7.5*10*55"
                oEdit.Text = "2"
            Case "10*10*55"
                oEdit.Text = "3"
            Case Else
                oEdit.Text = ""
     End Select

End Sub

Private Sub Impact_Size_Text_Select(oCob As ComboBox, oEdit As TextBox)
     
     Select Case Trim(oEdit.Text)
            Case "1"
                oCob.ListIndex = 1
            Case "2"
                oCob.ListIndex = 2
            Case "3"
                oCob.ListIndex = 3
            Case Else
                oCob.ListIndex = 0
     End Select

End Sub

Private Sub txt_SMP_NO_Change()
    txt_smp_no_p.Text = txt_SMP_NO.Text
    If Len(Trim(txt_smp_no_p.Text)) < 14 Then
        txt_SMP_CUT_LOC.Text = ""
    Else
        txt_SMP_CUT_LOC.Text = Find_SMP_LOC(Trim(txt_smp_no_p.Text))
    End If
End Sub

Private Sub TXT_TIM_IMPACT_SIZE_CD_Change()
    Call Impact_Size_Text_Select(Cob_TIM_IMPACT_SIZE, TXT_TIM_IMPACT_SIZE_CD)
End Sub

'Private Function Find_SMP_LOC(ByVal SMP_NO As String) As String
'
'    Set AdoRs = New adodb.Recordset
'
'    sQuery = "SELECT SMP_CUT_LOC FROM QP_TEST_HEAD WHERE SMP_NO = '" + Trim(SMP_NO) + "'"
'
'    AdoRs.Open sQuery, M_CN1, adOpenKeyset
'    If AdoRs.EOF Then
'        AdoRs.Close
'        Exit Function
'    End If
'    Find_SMP_LOC = AdoRs.Fields(0).Value
'    AdoRs.Close
'    Set AdoRs = Nothing
'
'End Function



