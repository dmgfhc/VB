VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AGB2010C 
   Caption         =   "轧制作业实绩查询及修改界面_AGB2010C"
   ClientHeight    =   9240
   ClientLeft      =   75
   ClientTop       =   1650
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   90
      TabIndex        =   43
      Top             =   75
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   1138
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_CB 
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
         Left            =   14025
         TabIndex        =   45
         Text            =   "CB"
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox TXT_UPD_EMP 
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
         Left            =   11415
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "作业人员"
         Top             =   180
         Width           =   1275
      End
      Begin VB.ComboBox CBO_GROUP 
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
         ItemData        =   "AGB2010C.frx":0000
         Left            =   9255
         List            =   "AGB2010C.frx":0010
         TabIndex        =   3
         Tag             =   "班别"
         Top             =   180
         Width           =   705
      End
      Begin VB.ComboBox CBO_SHIFT 
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
         ItemData        =   "AGB2010C.frx":0020
         Left            =   7380
         List            =   "AGB2010C.frx":002D
         TabIndex        =   2
         Tag             =   "班次"
         Top             =   180
         Width           =   705
      End
      Begin VB.ComboBox CBO_PLT 
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
         ItemData        =   "AGB2010C.frx":003A
         Left            =   5130
         List            =   "AGB2010C.frx":0044
         TabIndex        =   1
         Tag             =   "工厂代码"
         Top             =   180
         Width           =   705
      End
      Begin VB.ComboBox CBO_SLAB_NO 
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
         Left            =   1545
         TabIndex        =   0
         Tag             =   "板坯号"
         Top             =   165
         Width           =   1815
      End
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   330
         Top             =   165
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "板坯号"
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
         Left            =   3945
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "工厂代码"
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
         Left            =   6480
         Top             =   180
         Width           =   885
         _ExtentX        =   1561
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
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   8355
         Top             =   180
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         Caption         =   "班别"
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
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   10230
         Top             =   180
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8355
      Left            =   105
      TabIndex        =   80
      Top             =   840
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   14737
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
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
      TabCaption(0)   =   "轧制实绩"
      TabPicture(0)   =   "AGB2010C.frx":0050
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label65"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label61"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label59"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label49"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label43"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label28"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label14"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label15"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label16"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label17"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label18"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label19"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label20"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label22"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label23"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label24"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label25"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label26"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label27"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label30"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label31"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label32"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label5"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label4"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label12"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label21"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "ULabel18"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "ULabel90"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "SDB_MPLATE_CNT"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "ULabel4"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "ULabel3"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "ULabel13"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "ULabel12"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "SDB_COOL_WGT"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "SDB_COOL_ENT_TEMP"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "SDB_COOL_EXT_TEMP"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "SDB_COOL_AVE_TEMP"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "SDB_STAGE2_DEL_DUR"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "ULabel105"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "SDB_CR_STAGE3_THK"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "ULabel102"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "ULabel101"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "ULabel100"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "ULabel99"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "ULabel98"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "ULabel97"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "ULabel96"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "ULabel95"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "ULabel94"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "ULabel93"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "ULabel92"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "ULabel91"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "SDB_CR_STAGE3_TIME"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "SDB_CR_STAGE2_TIME"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "SDB_CR_STAGE1_TIME"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "SDB_CR_STAGE3_TEMP"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "SDB_CR_STAGE2_TEMP"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "SDB_CR_STAGE1_TEMP"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "SDB_CR_STAGE2_THK"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "SDB_CR_STAGE1_THK"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "ULabel80"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "TXT_MILL_END_TIME"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "TXT_MILL_STA_TIME"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "ULabel38"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "ULabel37"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "ULabel36"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "ULabel35"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "ULabel34"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "ULabel33"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "ULabel32"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "ULabel31"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "ULabel29"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "ULabel28"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "ULabel27"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "ULabel26"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "ULabel25"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "ULabel24"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "ULabel23"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "ULabel22"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "ULabel21"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "ULabel17"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "ULabel16"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "ULabel15"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "ULabel14"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "ULabel11"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "ULabel6"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "ULabel5"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "ULabel2"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "ULabel1"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "ULabel10"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "ULabel9"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "ULabel8"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "ULabel7"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "SDB_ROLLING_PASS"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "SDB_SLAB_MILL_LEN"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "SDB_TAIL_CROP_WID"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "SDB_TAIL_CROP_LEN"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "SDB_CROWN_VAL"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "SDB_MILL_END_MIN_TEMP"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "SDB_MILL_END_MAX_TEMP"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "SDB_END_AVE_TEMP"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "SDB_MILL_END_AIM_TEMP"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "SDB_HEAD_CROP_WID"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "SDB_HEAD_CROP_LEN"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "SDB_TAIL_WID"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "SDB_MID_WID"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "SDB_HEAD_WID"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "SDB_AVE_THK"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "SDB_AIM_THK"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "SDB_COILING_OUT_TEMP"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "SDB_COILING_IN_TEMP"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "SDB_TRIM_WGT"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "SDB_MAX_WID"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "SDB_MIN_WID"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "SDB_AVE_WID"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "SDB_AIM_WID"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "SDB_DELAY_DUR"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "SDB_MILL_DUR"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "ULabel81"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "ULabel89"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "SDB_STD_WID"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "SDB_HD_CUT_THK"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "CHK_NON_CR_CD"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "CHK_CR_CD"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "TXT_CR_CD"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "TXT_ROLLING_METHOD"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "CHK_ROLLING_AUTO"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "CHK_ROLLING_OP"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "CHK_COOL_PRTY_COIL"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "CHK_COOL_PRTY_FIN"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "TXT_COOL_PRTY"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "CHK_COOL_METHOD_WATER"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "CHK_COOL_METHOD_AIR"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "CHK_PROD_COIL"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "CHK_PROD_PLATE"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "TXT_COOLING_METHOD"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "TXT_PROD_STATUS"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "ULabel41"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "CHK_PROD_C_PLATE"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "CHK_PROD_C_COIL"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "ROLL_EMP"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "ROLL_GROUP"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "ROLL_SHIFT"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).ControlCount=   151
      TabCaption(1)   =   "母板实绩"
      TabPicture(1)   =   "AGB2010C.frx":006C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label114"
      Tab(1).Control(1)=   "Label96"
      Tab(1).Control(2)=   "SDB_MOTHER_SCH_LEN12"
      Tab(1).Control(3)=   "SDB_MOTHER_SCH_LEN11"
      Tab(1).Control(4)=   "SDB_MOTHER_SCH_LEN10"
      Tab(1).Control(5)=   "SDB_MOTHER_SCH_LEN9"
      Tab(1).Control(6)=   "SDB_MOTHER_SCH_LEN8"
      Tab(1).Control(7)=   "SDB_MOTHER_SCH_LEN7"
      Tab(1).Control(8)=   "SDB_MOTHER_SCH_LEN6"
      Tab(1).Control(9)=   "SDB_MOTHER_SCH_LEN5"
      Tab(1).Control(10)=   "SDB_MOTHER_SCH_LEN4"
      Tab(1).Control(11)=   "SDB_MOTHER_SCH_LEN3"
      Tab(1).Control(12)=   "SDB_MOTHER_SCH_LEN2"
      Tab(1).Control(13)=   "SDB_MOTHER_SCH_LEN1"
      Tab(1).Control(14)=   "SSCommand2"
      Tab(1).Control(15)=   "SSCommand1"
      Tab(1).Control(16)=   "ULabel74"
      Tab(1).Control(17)=   "ULabel73"
      Tab(1).Control(18)=   "ULabel72"
      Tab(1).Control(19)=   "ULabel71"
      Tab(1).Control(20)=   "ULabel70"
      Tab(1).Control(21)=   "ULabel69"
      Tab(1).Control(22)=   "ULabel68"
      Tab(1).Control(23)=   "ULabel67"
      Tab(1).Control(24)=   "ULabel66"
      Tab(1).Control(25)=   "ULabel65"
      Tab(1).Control(26)=   "ULabel64"
      Tab(1).Control(27)=   "ULabel63"
      Tab(1).Control(28)=   "ULabel62"
      Tab(1).Control(29)=   "ULabel61"
      Tab(1).Control(30)=   "ULabel60"
      Tab(1).Control(31)=   "ULabel59"
      Tab(1).Control(32)=   "ULabel58"
      Tab(1).Control(33)=   "ULabel57"
      Tab(1).Control(34)=   "ULabel56"
      Tab(1).Control(35)=   "ULabel55"
      Tab(1).Control(36)=   "ULabel54"
      Tab(1).Control(37)=   "ULabel53"
      Tab(1).Control(38)=   "ULabel52"
      Tab(1).Control(39)=   "ULabel51"
      Tab(1).Control(40)=   "SDB_MOTHER_PLATE_LEN12"
      Tab(1).Control(41)=   "SDB_MOTHER_PLATE_LEN11"
      Tab(1).Control(42)=   "SDB_MOTHER_PLATE_LEN10"
      Tab(1).Control(43)=   "SDB_MOTHER_PLATE_LEN9"
      Tab(1).Control(44)=   "SDB_MOTHER_PLATE_LEN8"
      Tab(1).Control(45)=   "SDB_MOTHER_PLATE_LEN7"
      Tab(1).Control(46)=   "SDB_MOTHER_PLATE_LEN6"
      Tab(1).Control(47)=   "SDB_MOTHER_PLATE_LEN5"
      Tab(1).Control(48)=   "SDB_MOTHER_PLATE_LEN4"
      Tab(1).Control(49)=   "SDB_MOTHER_PLATE_LEN3"
      Tab(1).Control(50)=   "SDB_MOTHER_PLATE_LEN2"
      Tab(1).Control(51)=   "SDB_MOTHER_PLATE_LEN1"
      Tab(1).Control(52)=   "TXT_MOTHER_PLATE10"
      Tab(1).Control(53)=   "TXT_MOTHER_PLATE12"
      Tab(1).Control(54)=   "TXT_MOTHER_PLATE8"
      Tab(1).Control(55)=   "TXT_MOTHER_PLATE9"
      Tab(1).Control(56)=   "TXT_MOTHER_PLATE11"
      Tab(1).Control(57)=   "TXT_MOTHER_PLATE7"
      Tab(1).Control(58)=   "TXT_MOTHER_PLATE4"
      Tab(1).Control(59)=   "TXT_MOTHER_PLATE6"
      Tab(1).Control(60)=   "TXT_MOTHER_PLATE2"
      Tab(1).Control(61)=   "TXT_MOTHER_PLATE3"
      Tab(1).Control(62)=   "TXT_MOTHER_PLATE5"
      Tab(1).Control(63)=   "TXT_MOTHER_PLATE1"
      Tab(1).Control(64)=   "TXT_COMFRM"
      Tab(1).Control(65)=   "CHK_YES"
      Tab(1).Control(66)=   "CHK_NO"
      Tab(1).Control(67)=   "TXT_CUTEND_CD"
      Tab(1).ControlCount=   68
      Begin VB.TextBox ROLL_SHIFT 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8430
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   6495
         Width           =   555
      End
      Begin VB.TextBox ROLL_GROUP 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   149
         Top             =   6495
         Width           =   555
      End
      Begin VB.TextBox ROLL_EMP 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11670
         Locked          =   -1  'True
         TabIndex        =   148
         Top             =   6495
         Width           =   1005
      End
      Begin VB.TextBox TXT_CUTEND_CD 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69360
         TabIndex        =   134
         Top             =   765
         Width           =   540
      End
      Begin VB.CheckBox CHK_NO 
         BackColor       =   &H00FFFF80&
         Caption         =   "余材"
         Height          =   240
         Left            =   -72375
         TabIndex        =   123
         Top             =   1035
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CheckBox CHK_YES 
         BackColor       =   &H00FFFF80&
         Caption         =   "订单"
         Height          =   240
         Left            =   -72375
         TabIndex        =   122
         Top             =   765
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox TXT_COMFRM 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73020
         TabIndex        =   121
         Top             =   765
         Width           =   540
      End
      Begin VB.CheckBox CHK_PROD_C_COIL 
         Caption         =   "卷轧钢卷"
         Height          =   240
         Left            =   6165
         TabIndex        =   120
         Top             =   5370
         Width           =   1110
      End
      Begin VB.CheckBox CHK_PROD_C_PLATE 
         Caption         =   "卷轧钢板"
         Height          =   240
         Left            =   6165
         TabIndex        =   119
         Top             =   4875
         Width           =   1110
      End
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   330
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "控轧空闲时间"
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
      Begin VB.TextBox TXT_MOTHER_PLATE1 
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
         Left            =   -73020
         MaxLength       =   2
         TabIndex        =   56
         Text            =   "01"
         Top             =   1665
         Width           =   570
      End
      Begin VB.TextBox TXT_MOTHER_PLATE5 
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
         Left            =   -73020
         MaxLength       =   2
         TabIndex        =   64
         Text            =   "05"
         Top             =   4785
         Width           =   570
      End
      Begin VB.TextBox TXT_MOTHER_PLATE3 
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
         Left            =   -73020
         MaxLength       =   2
         TabIndex        =   60
         Text            =   "03"
         Top             =   3225
         Width           =   570
      End
      Begin VB.TextBox TXT_MOTHER_PLATE2 
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
         Left            =   -73020
         MaxLength       =   2
         TabIndex        =   58
         Text            =   "02"
         Top             =   2505
         Width           =   570
      End
      Begin VB.TextBox TXT_MOTHER_PLATE6 
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
         Left            =   -73020
         MaxLength       =   2
         TabIndex        =   66
         Text            =   "06"
         Top             =   5625
         Width           =   570
      End
      Begin VB.TextBox TXT_MOTHER_PLATE4 
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
         Left            =   -73020
         MaxLength       =   2
         TabIndex        =   62
         Text            =   "04"
         Top             =   4065
         Width           =   570
      End
      Begin VB.TextBox TXT_MOTHER_PLATE7 
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
         Left            =   -65835
         MaxLength       =   2
         TabIndex        =   68
         Text            =   "07"
         Top             =   1665
         Width           =   540
      End
      Begin VB.TextBox TXT_MOTHER_PLATE11 
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
         Left            =   -65835
         MaxLength       =   2
         TabIndex        =   76
         Text            =   "11"
         Top             =   4785
         Width           =   540
      End
      Begin VB.TextBox TXT_MOTHER_PLATE9 
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
         Left            =   -65835
         MaxLength       =   2
         TabIndex        =   72
         Text            =   "09"
         Top             =   3225
         Width           =   540
      End
      Begin VB.TextBox TXT_MOTHER_PLATE8 
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
         Left            =   -65835
         MaxLength       =   2
         TabIndex        =   70
         Text            =   "08"
         Top             =   2505
         Width           =   540
      End
      Begin VB.TextBox TXT_MOTHER_PLATE12 
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
         Left            =   -65835
         MaxLength       =   2
         TabIndex        =   78
         Text            =   "12"
         Top             =   5625
         Width           =   540
      End
      Begin VB.TextBox TXT_MOTHER_PLATE10 
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
         Left            =   -65835
         MaxLength       =   2
         TabIndex        =   74
         Text            =   "10"
         Top             =   4065
         Width           =   540
      End
      Begin VB.TextBox TXT_PROD_STATUS 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   85
         Tag             =   "生产类型"
         Text            =   " "
         Top             =   4875
         Width           =   495
      End
      Begin VB.TextBox TXT_COOLING_METHOD 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   84
         Tag             =   "冷却方式"
         Text            =   " "
         Top             =   3510
         Width           =   495
      End
      Begin VB.CheckBox CHK_PROD_PLATE 
         Caption         =   "平轧钢板"
         Height          =   240
         Left            =   6165
         TabIndex        =   10
         Top             =   4605
         Width           =   1110
      End
      Begin VB.CheckBox CHK_PROD_COIL 
         Caption         =   "平轧钢卷"
         Height          =   240
         Left            =   6165
         TabIndex        =   11
         Top             =   5115
         Width           =   1110
      End
      Begin VB.CheckBox CHK_COOL_METHOD_AIR 
         Caption         =   "空冷"
         Height          =   240
         Left            =   2355
         TabIndex        =   12
         Top             =   3480
         Width           =   960
      End
      Begin VB.CheckBox CHK_COOL_METHOD_WATER 
         Caption         =   "水冷"
         Height          =   195
         Left            =   2355
         TabIndex        =   13
         Top             =   3750
         Width           =   960
      End
      Begin VB.TextBox TXT_COOL_PRTY 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   83
         Tag             =   "温度代码"
         Text            =   " "
         Top             =   4290
         Width           =   495
      End
      Begin VB.CheckBox CHK_COOL_PRTY_FIN 
         Caption         =   "终轧温度"
         Height          =   240
         Left            =   2355
         TabIndex        =   24
         Top             =   4290
         Width           =   1095
      End
      Begin VB.CheckBox CHK_COOL_PRTY_COIL 
         Caption         =   "冷却出口温度"
         Height          =   195
         Left            =   2355
         TabIndex        =   25
         Top             =   4560
         Width           =   1395
      End
      Begin VB.CheckBox CHK_ROLLING_OP 
         Caption         =   "人工干预"
         Height          =   285
         Left            =   6165
         TabIndex        =   21
         Top             =   4110
         Width           =   1155
      End
      Begin VB.CheckBox CHK_ROLLING_AUTO 
         Caption         =   "自动"
         Height          =   285
         Left            =   6165
         TabIndex        =   20
         Top             =   3840
         Width           =   1155
      End
      Begin VB.TextBox TXT_ROLLING_METHOD 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   82
         Tag             =   "轧制方式"
         Text            =   " "
         Top             =   3945
         Width           =   495
      End
      Begin VB.TextBox TXT_CR_CD 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9100
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   81
         Tag             =   "控轧代码"
         Text            =   " "
         Top             =   5085
         Width           =   555
      End
      Begin VB.CheckBox CHK_CR_CD 
         Caption         =   "控轧"
         Height          =   285
         Left            =   9690
         TabIndex        =   44
         Top             =   5040
         Width           =   915
      End
      Begin VB.CheckBox CHK_NON_CR_CD 
         Caption         =   "否"
         Height          =   285
         Left            =   9690
         TabIndex        =   46
         Top             =   5310
         Width           =   915
      End
      Begin CSTextLibCtl.sidbEdit SDB_HD_CUT_THK 
         Height          =   330
         Left            =   9105
         TabIndex        =   28
         Tag             =   "标准厚度"
         Top             =   1785
         Width           =   1230
         _Version        =   262145
         _ExtentX        =   2170
         _ExtentY        =   582
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
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_STD_WID 
         Height          =   330
         Left            =   5610
         TabIndex        =   16
         Tag             =   " 标准宽度"
         Top             =   1785
         Width           =   1230
         _Version        =   262145
         _ExtentX        =   2170
         _ExtentY        =   582
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel89 
         Height          =   330
         Left            =   7620
         Top             =   1785
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "标准厚度"
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
      Begin InDate.ULabel ULabel81 
         Height          =   315
         Left            =   4125
         Top             =   1785
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   " 标准宽度"
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
      Begin CSTextLibCtl.sidbEdit SDB_MILL_DUR 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   2715
         Width           =   1200
         _Version        =   262145
         _ExtentX        =   2117
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
      Begin CSTextLibCtl.sidbEdit SDB_DELAY_DUR 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1785
         Width           =   1200
         _Version        =   262145
         _ExtentX        =   2117
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
         MaxValue        =   900
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_AIM_WID 
         Height          =   315
         Left            =   5610
         TabIndex        =   14
         Top             =   870
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_AVE_WID 
         Height          =   315
         Left            =   5610
         TabIndex        =   15
         Tag             =   "平均宽度"
         Top             =   1320
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MIN_WID 
         Height          =   315
         Left            =   5610
         TabIndex        =   17
         Top             =   2250
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MAX_WID 
         Height          =   315
         Left            =   5610
         TabIndex        =   18
         Top             =   2715
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_TRIM_WGT 
         Height          =   315
         Left            =   5610
         TabIndex        =   19
         Top             =   3180
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_COILING_IN_TEMP 
         Height          =   315
         Left            =   5610
         TabIndex        =   22
         Top             =   5955
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_COILING_OUT_TEMP 
         Height          =   315
         Left            =   5610
         TabIndex        =   23
         Top             =   6435
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_AIM_THK 
         Height          =   315
         Left            =   9105
         TabIndex        =   26
         Top             =   870
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_AVE_THK 
         Height          =   315
         Left            =   9105
         TabIndex        =   27
         Tag             =   "平均厚度"
         Top             =   1320
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_HEAD_WID 
         Height          =   315
         Left            =   9105
         TabIndex        =   29
         Top             =   2250
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MID_WID 
         Height          =   315
         Left            =   9105
         TabIndex        =   30
         Top             =   2715
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_TAIL_WID 
         Height          =   315
         Left            =   9105
         TabIndex        =   31
         Top             =   3180
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_HEAD_CROP_LEN 
         Height          =   315
         Left            =   9105
         TabIndex        =   33
         Top             =   4155
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_HEAD_CROP_WID 
         Height          =   315
         Left            =   9105
         TabIndex        =   34
         Top             =   4500
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MILL_END_AIM_TEMP 
         Height          =   315
         Left            =   12735
         TabIndex        =   35
         Top             =   870
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_END_AVE_TEMP 
         Height          =   315
         Left            =   12735
         TabIndex        =   36
         Tag             =   "终轧平均温度"
         Top             =   1320
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MILL_END_MAX_TEMP 
         Height          =   315
         Left            =   12735
         TabIndex        =   37
         Top             =   1785
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MILL_END_MIN_TEMP 
         Height          =   315
         Left            =   12735
         TabIndex        =   38
         Top             =   2250
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CROWN_VAL 
         Height          =   315
         Left            =   12735
         TabIndex        =   39
         Top             =   2715
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_TAIL_CROP_LEN 
         Height          =   315
         Left            =   12735
         TabIndex        =   41
         Top             =   4155
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_TAIL_CROP_WID 
         Height          =   315
         Left            =   12735
         TabIndex        =   42
         Top             =   4500
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_SLAB_MILL_LEN 
         Height          =   315
         Left            =   9105
         TabIndex        =   32
         Tag             =   "轧制长度"
         Top             =   3660
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_ROLLING_PASS 
         Height          =   315
         Left            =   12735
         TabIndex        =   40
         Top             =   3180
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   330
         Top             =   870
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "开轧时间"
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
         Left            =   330
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "终轧时间"
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
         Left            =   330
         Top             =   1770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "轧制空闲时间"
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   330
         Top             =   2715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "纯轧时间"
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
         Left            =   4125
         Top             =   4875
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "生产类型"
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
         Left            =   330
         Top             =   3510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "冷却方式"
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   4125
         Top             =   870
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "目标宽度"
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   4125
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "平均宽度"
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   4125
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "最小宽度"
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
         Left            =   4125
         Top             =   2715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "最大宽度"
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   4125
         Top             =   5955
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "入口卷取炉温度"
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   4125
         Top             =   6420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "出口卷取炉温度"
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
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   4125
         Top             =   3180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "立辊减宽量"
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
         Left            =   330
         Top             =   4290
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "冷却温度代码"
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
         Left            =   7620
         Top             =   870
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "目标厚度"
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   7620
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "平均厚度"
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
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   7620
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "头部宽度"
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
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   7620
         Top             =   2715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "中部宽度"
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
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   7620
         Top             =   4155
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "切头：长度"
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
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   7620
         Top             =   4500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "     宽度"
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
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   7620
         Top             =   3180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "尾部宽度"
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
         Left            =   7620
         Top             =   3660
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "轧制长度"
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
         Left            =   11250
         Top             =   870
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "终轧目标温度"
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
         Left            =   11250
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "终轧平均温度"
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
         Left            =   11250
         Top             =   1785
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "终轧最高温度"
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
         Left            =   11250
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "终轧最低温度"
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
         Left            =   11250
         Top             =   4155
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "切尾：长度"
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
         Height          =   315
         Left            =   11250
         Top             =   4500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "     宽度"
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   11250
         Top             =   2715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "凸度"
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
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   11250
         Top             =   3180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "道次数"
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
      Begin CSTextLibCtl.sitxEdit TXT_MILL_STA_TIME 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Tag             =   "开轧时间"
         Top             =   870
         Width           =   2145
         _Version        =   262145
         _ExtentX        =   3784
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__:__ "
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
         Mask            =   "____-__-__ __:__:__ "
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin CSTextLibCtl.sitxEdit TXT_MILL_END_TIME 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Tag             =   "终轧时间"
         Top             =   1320
         Width           =   2145
         _Version        =   262145
         _ExtentX        =   3784
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
      Begin InDate.ULabel ULabel80 
         Height          =   315
         Left            =   4125
         Top             =   3945
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "轧制方式"
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
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE1_THK 
         Height          =   315
         Left            =   9855
         TabIndex        =   47
         Tag             =   "一阶段厚度"
         Top             =   5715
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE2_THK 
         Height          =   315
         Left            =   9855
         TabIndex        =   50
         Tag             =   "二阶段厚度"
         Top             =   7215
         Visible         =   0   'False
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE1_TEMP 
         Height          =   315
         Left            =   11700
         TabIndex        =   48
         Tag             =   "一阶段温度"
         Top             =   5715
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE2_TEMP 
         Height          =   315
         Left            =   11700
         TabIndex        =   51
         Tag             =   "二阶段温度"
         Top             =   7215
         Visible         =   0   'False
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE3_TEMP 
         Height          =   315
         Left            =   11700
         TabIndex        =   54
         Tag             =   "三阶段温度"
         Top             =   7665
         Visible         =   0   'False
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE1_TIME 
         Height          =   315
         Left            =   13725
         TabIndex        =   49
         Tag             =   "待轧时间一阶段"
         Top             =   5715
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE2_TIME 
         Height          =   315
         Left            =   13725
         TabIndex        =   52
         Tag             =   "二阶段待轧时间"
         Top             =   7215
         Visible         =   0   'False
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE3_TIME 
         Height          =   315
         Left            =   13725
         TabIndex        =   55
         Tag             =   "三阶段待轧时间"
         Top             =   7665
         Visible         =   0   'False
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel91 
         Height          =   315
         Left            =   9105
         Top             =   5715
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Caption         =   "厚度"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel92 
         Height          =   315
         Left            =   9105
         Top             =   7215
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Caption         =   "厚度"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel93 
         Height          =   315
         Left            =   9105
         Top             =   7665
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Caption         =   "厚度"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel94 
         Height          =   315
         Left            =   10980
         Top             =   5715
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Caption         =   "温度"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel95 
         Height          =   315
         Left            =   10980
         Top             =   7215
         Visible         =   0   'False
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Caption         =   "温度"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel96 
         Height          =   315
         Left            =   10980
         Top             =   7665
         Visible         =   0   'False
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Caption         =   "温度"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel97 
         Height          =   315
         Left            =   12750
         Top             =   5715
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "待轧时间"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel98 
         Height          =   315
         Left            =   12750
         Top             =   7215
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "待轧时间"
         Alignment       =   1
         BackColor       =   14804174
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
         Height          =   315
         Left            =   12750
         Top             =   7665
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "待轧时间"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel100 
         Height          =   315
         Left            =   7620
         Top             =   5715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "一阶段"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel101 
         Height          =   315
         Left            =   7620
         Top             =   7215
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "二阶段"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin InDate.ULabel ULabel102 
         Height          =   315
         Left            =   7620
         Top             =   7665
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "三阶段"
         Alignment       =   1
         BackColor       =   14804174
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
      Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE3_THK 
         Height          =   315
         Left            =   9855
         TabIndex        =   53
         Tag             =   "三阶段厚度"
         Top             =   7665
         Visible         =   0   'False
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         MaxValue        =   999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel105 
         Height          =   315
         Left            =   7620
         Top             =   5085
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "控轧代码"
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN1 
         Height          =   315
         Left            =   -70935
         TabIndex        =   57
         Top             =   1665
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN2 
         Height          =   315
         Left            =   -70935
         TabIndex        =   59
         Top             =   2505
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN3 
         Height          =   315
         Left            =   -70935
         TabIndex        =   61
         Top             =   3225
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN4 
         Height          =   315
         Left            =   -70935
         TabIndex        =   63
         Top             =   4065
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN5 
         Height          =   315
         Left            =   -70935
         TabIndex        =   65
         Top             =   4785
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN6 
         Height          =   315
         Left            =   -70935
         TabIndex        =   67
         Top             =   5625
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN7 
         Height          =   315
         Left            =   -63780
         TabIndex        =   69
         Top             =   1665
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN8 
         Height          =   315
         Left            =   -63780
         TabIndex        =   71
         Top             =   2505
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN9 
         Height          =   315
         Left            =   -63780
         TabIndex        =   73
         Top             =   3225
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN10 
         Height          =   315
         Left            =   -63780
         TabIndex        =   75
         Top             =   4065
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN11 
         Height          =   315
         Left            =   -63780
         TabIndex        =   77
         Top             =   4785
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN12 
         Height          =   315
         Left            =   -63780
         TabIndex        =   79
         Top             =   5625
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
      Begin InDate.ULabel ULabel51 
         Height          =   315
         Left            =   -74505
         Top             =   1665
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  1"
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
      Begin InDate.ULabel ULabel52 
         Height          =   315
         Left            =   -74505
         Top             =   2505
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  2"
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
      Begin InDate.ULabel ULabel53 
         Height          =   315
         Left            =   -74505
         Top             =   3225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  3"
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
      Begin InDate.ULabel ULabel54 
         Height          =   315
         Left            =   -74505
         Top             =   4065
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  4"
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
         Left            =   -74505
         Top             =   4785
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  5"
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
      Begin InDate.ULabel ULabel56 
         Height          =   315
         Left            =   -74505
         Top             =   5625
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  6"
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
      Begin InDate.ULabel ULabel57 
         Height          =   315
         Left            =   -72420
         Top             =   1665
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel58 
         Height          =   315
         Left            =   -72420
         Top             =   2505
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel59 
         Height          =   315
         Left            =   -72420
         Top             =   3225
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel60 
         Height          =   315
         Left            =   -72420
         Top             =   4065
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel61 
         Height          =   315
         Left            =   -72420
         Top             =   4785
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel62 
         Height          =   315
         Left            =   -72420
         Top             =   5625
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel63 
         Height          =   315
         Left            =   -67320
         Top             =   1665
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  7"
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
      Begin InDate.ULabel ULabel64 
         Height          =   315
         Left            =   -67320
         Top             =   2505
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  8"
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
      Begin InDate.ULabel ULabel65 
         Height          =   315
         Left            =   -67320
         Top             =   3225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  9"
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
      Begin InDate.ULabel ULabel66 
         Height          =   315
         Left            =   -67320
         Top             =   4065
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  10"
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
      Begin InDate.ULabel ULabel67 
         Height          =   315
         Left            =   -67320
         Top             =   4785
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  11"
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
      Begin InDate.ULabel ULabel68 
         Height          =   315
         Left            =   -67320
         Top             =   5625
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板  12"
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
      Begin InDate.ULabel ULabel69 
         Height          =   315
         Left            =   -65265
         Top             =   1665
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel70 
         Height          =   315
         Left            =   -65265
         Top             =   2505
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel71 
         Height          =   315
         Left            =   -65265
         Top             =   3225
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel72 
         Height          =   315
         Left            =   -65265
         Top             =   4065
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel73 
         Height          =   315
         Left            =   -65265
         Top             =   4785
         Width           =   1455
         _ExtentX        =   2566
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
      Begin InDate.ULabel ULabel74 
         Height          =   315
         Left            =   -65265
         Top             =   5625
         Width           =   1455
         _ExtentX        =   2566
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
      Begin CSTextLibCtl.sidbEdit SDB_STAGE2_DEL_DUR 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   2250
         Width           =   1200
         _Version        =   262145
         _ExtentX        =   2117
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
         MaxValue        =   900
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_COOL_AVE_TEMP 
         Height          =   315
         Left            =   1800
         TabIndex        =   124
         Tag             =   "冷却平均温度"
         Top             =   5040
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_COOL_EXT_TEMP 
         Height          =   315
         Left            =   1800
         TabIndex        =   125
         Top             =   5955
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_COOL_ENT_TEMP 
         Height          =   315
         Left            =   1800
         TabIndex        =   126
         Top             =   5505
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_COOL_WGT 
         Height          =   315
         Left            =   1800
         TabIndex        =   127
         Tag             =   "冷却水量"
         Top             =   6435
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   330
         Top             =   5040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "冷却平均温度"
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
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   330
         Top             =   5955
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "冷却出口温度"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   330
         Top             =   5505
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "冷却入口温度"
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
         Left            =   330
         Top             =   6435
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "冷却水量"
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
      Begin CSTextLibCtl.sidbEdit SDB_MPLATE_CNT 
         Height          =   330
         Left            =   12735
         TabIndex        =   128
         Tag             =   "母板数量"
         Top             =   5085
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   582
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
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel90 
         Height          =   315
         Left            =   11250
         Top             =   5085
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "母板数量"
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
      Begin Threed.SSCommand SSCommand1 
         Height          =   330
         Left            =   -74490
         TabIndex        =   133
         Top             =   765
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   14737632
         BackStyle       =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "缺号母板确定"
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   330
         Left            =   -71415
         TabIndex        =   135
         Top             =   765
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   14737632
         BackStyle       =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "母板剪切结束确定"
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN1 
         Height          =   315
         Left            =   -69720
         TabIndex        =   136
         Top             =   1665
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN2 
         Height          =   315
         Left            =   -69720
         TabIndex        =   137
         Top             =   2505
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN3 
         Height          =   315
         Left            =   -69720
         TabIndex        =   138
         Top             =   3225
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN4 
         Height          =   315
         Left            =   -69720
         TabIndex        =   139
         Top             =   4065
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN5 
         Height          =   315
         Left            =   -69720
         TabIndex        =   140
         Top             =   4785
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN6 
         Height          =   315
         Left            =   -69720
         TabIndex        =   141
         Top             =   5625
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN7 
         Height          =   315
         Left            =   -62550
         TabIndex        =   142
         Top             =   1665
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN8 
         Height          =   315
         Left            =   -62550
         TabIndex        =   143
         Top             =   2505
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN9 
         Height          =   315
         Left            =   -62550
         TabIndex        =   144
         Top             =   3225
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN10 
         Height          =   315
         Left            =   -62550
         TabIndex        =   145
         Top             =   4065
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN11 
         Height          =   315
         Left            =   -62550
         TabIndex        =   146
         Top             =   4785
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN12 
         Height          =   315
         Left            =   -62550
         TabIndex        =   147
         Top             =   5625
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         ReadOnly        =   -1  'True
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   7650
         Top             =   6435
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         Caption         =   "  班次(      )  班别(      )  作业人员(          )"
         Alignment       =   0
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
      Begin VB.Label Label21 
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3060
         TabIndex        =   132
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2955
         TabIndex        =   131
         Top             =   6435
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3060
         TabIndex        =   130
         Top             =   5955
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3060
         TabIndex        =   129
         Top             =   5505
         Width           =   255
      End
      Begin VB.Label Label32 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14025
         TabIndex        =   118
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13995
         TabIndex        =   117
         Top             =   2250
         Width           =   255
      End
      Begin VB.Label Label30 
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13995
         TabIndex        =   116
         Top             =   1785
         Width           =   255
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   24945
         TabIndex        =   115
         Top             =   -2880
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13995
         TabIndex        =   114
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   113
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   112
         Top             =   3660
         Width           =   255
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   111
         Top             =   3180
         Width           =   255
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   110
         Top             =   2715
         Width           =   255
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   109
         Top             =   2250
         Width           =   255
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   108
         Top             =   1785
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   107
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   106
         Top             =   2715
         Width           =   315
      End
      Begin VB.Label Label17 
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   105
         Top             =   2250
         Width           =   315
      End
      Begin VB.Label Label16 
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   104
         Top             =   1785
         Width           =   345
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6855
         TabIndex        =   103
         Top             =   3180
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6855
         TabIndex        =   102
         Top             =   2715
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6855
         TabIndex        =   101
         Top             =   2250
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6855
         TabIndex        =   100
         Top             =   1785
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6855
         TabIndex        =   99
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6165
         TabIndex        =   98
         Top             =   3885
         Width           =   510
      End
      Begin VB.Label Label96 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68325
         TabIndex        =   97
         Top             =   1665
         Width           =   255
      End
      Begin VB.Label Label114 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -61215
         TabIndex        =   96
         Top             =   1665
         Width           =   255
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6825
         TabIndex        =   95
         Top             =   5955
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6855
         TabIndex        =   94
         Top             =   870
         Width           =   255
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   93
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10380
         TabIndex        =   92
         Top             =   870
         Width           =   255
      End
      Begin VB.Label Label59 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14025
         TabIndex        =   91
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label61 
         Caption         =   "μm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13995
         TabIndex        =   90
         Top             =   2715
         Width           =   375
      End
      Begin VB.Label Label65 
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13995
         TabIndex        =   89
         Top             =   870
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10515
         TabIndex        =   88
         Top             =   5715
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12300
         TabIndex        =   87
         Top             =   5715
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "min"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14355
         TabIndex        =   86
         Top             =   5715
         Width           =   375
      End
   End
End
Attribute VB_Name = "AGB2010C"
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
'-- Program Name      轧制作业实绩查询及修改界面
'-- Program ID        AGB2010C
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
Public sQuery_Rt As String          'Active Form sQuery Setting
       
Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

Dim pControl1 As New Collection     'Master Primary Key Collection
Dim nControl1 As New Collection     'Master Necessary Collection
Dim mControl1 As New Collection     'Master Maxlength check Collection
Dim iControl1 As New Collection     'Master Insert Collection
Dim rControl1 As New Collection     'Master Refer Collection
Dim cControl1 As New Collection     'Master Copy Collection
Dim aControl1 As New Collection     'Master -> Spread Collection
Dim lControl1 As New Collection     'Master Lock Collection

Dim sControl  As New Collection     'Master Clear Key Collection
Dim MC        As New Collection     'Master Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(CBO_SLAB_NO, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_group, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_UPD_EMP, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_MILL_STA_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_MILL_END_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_DELAY_DUR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_STAGE2_DEL_DUR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_MILL_DUR, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_PROD_STATUS, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_COOLING_METHOD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_COOL_AVE_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_COOL_EXT_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_COOL_ENT_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_COOL_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_AIM_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_AVE_WID, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_STD_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_MIN_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_MAX_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_TRIM_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ROLLING_METHOD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_COILING_IN_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDB_COILING_OUT_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_COOL_PRTY, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_AIM_THK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_AVE_THK, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_HD_CUT_THK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '热卡量厚度
          Call Gp_Ms_Collection(SDB_HEAD_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_MID_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_TAIL_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_SLAB_MILL_LEN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_HEAD_CROP_LEN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_HEAD_CROP_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDB_MILL_END_AIM_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_END_AVE_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDB_MILL_END_MAX_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDB_MILL_END_MIN_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_CROWN_VAL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ROLLING_PASS, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_TAIL_CROP_LEN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_TAIL_CROP_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_CR_CD, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_CR_STAGE1_THK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_CR_STAGE1_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_CR_STAGE1_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_CR_STAGE2_THK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_CR_STAGE2_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_CR_STAGE2_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(ROLL_SHIFT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(ROLL_GROUP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(ROLL_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_MPLATE_CNT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_CB, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
        
           Call Gp_Ms_Collection(CBO_SLAB_NO, "p", "n", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          '     Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(cbo_group, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     '      Call Gp_Ms_Collection(TXT_UPD_EMP, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE1, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN1, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE2, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE3, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN3, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE4, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN4, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN4, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE5, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN5, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN5, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE6, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN6, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN6, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE7, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN7, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN7, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE8, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN8, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN8, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE9, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN9, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN9, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(TXT_MOTHER_PLATE10, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN10, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN10, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(TXT_MOTHER_PLATE11, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN11, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN11, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(TXT_MOTHER_PLATE12, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN12, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN12, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(TXT_COMFRM, "p", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_CUTEND_CD, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
                                
     Call Gp_Clear_Collection(CHK_COOL_METHOD_AIR, "s", sControl)
     Call Gp_Clear_Collection(CHK_COOL_METHOD_WATER, "s", sControl)
     Call Gp_Clear_Collection(CHK_COOL_PRTY_FIN, "s", sControl)
     Call Gp_Clear_Collection(CHK_COOL_PRTY_COIL, "s", sControl)
     Call Gp_Clear_Collection(CHK_ROLLING_AUTO, "s", sControl)
     Call Gp_Clear_Collection(CHK_ROLLING_OP, "s", sControl)
     Call Gp_Clear_Collection(CHK_PROD_PLATE, "s", sControl)
     Call Gp_Clear_Collection(CHK_PROD_C_PLATE, "s", cControl)
     Call Gp_Clear_Collection(CHK_PROD_COIL, "s", sControl)
     Call Gp_Clear_Collection(CHK_PROD_C_COIL, "s", sControl)
     Call Gp_Clear_Collection(CHK_CR_CD, "s", sControl)
     Call Gp_Clear_Collection(CHK_NON_CR_CD, "s", sControl)
     
     MC.Add Item:=sControl, Key:="sControl"

    'MASTER Collection
     Mc1.Add Item:="AGB2010C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="AGB2010C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Mc2.Add Item:="AGB2010C.P_MODIFY2", Key:="P-M"
     Mc2.Add Item:="AGB2010C.P_REFER2", Key:="P-R"
     Mc2.Add Item:=pControl1, Key:="pControl"
     Mc2.Add Item:=nControl1, Key:="nControl"
     Mc2.Add Item:=mControl1, Key:="mControl"
     Mc2.Add Item:=iControl1, Key:="iControl"
     Mc2.Add Item:=rControl1, Key:="rControl"
     Mc2.Add Item:=cControl1, Key:="cControl"
     Mc2.Add Item:=aControl1, Key:="aControl"
     Mc2.Add Item:=lControl1, Key:="lControl"
     
     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub

Private Sub CBO_SLAB_NO_Change()
   Dim SMESG As String
      If Len(CBO_SLAB_NO.Text) > 10 Then
      SMESG = "板坯号长度不能超过10位，请确认板坯号 ！！！"
      Call Gp_MsgBoxDisplay(SMESG)
   End If
End Sub

Private Sub CBO_SLAB_NO_Click()
    CBO_SLAB_NO.Text = Mid(CBO_SLAB_NO.Text, 1, 10)
'    Call Form_Ref
End Sub

Private Sub Chk_Cool_Method_Air_Click()
   
    If CHK_COOL_METHOD_AIR.Value = ssCBUnchecked Then
       If CHK_COOL_METHOD_WATER.Value = ssCBUnchecked Then
         ' CHK_COOL_METHOD_AIR.Value = ssCBChecked
          TXT_COOLING_METHOD.Text = ""
          CHK_COOL_METHOD_AIR.ForeColor = &H80000012
          CHK_COOL_METHOD_WATER.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_COOLING_METHOD.Text = "A"
   
   CHK_COOL_METHOD_AIR.ForeColor = &HFF&
   CHK_COOL_METHOD_AIR.Value = ssCBChecked

   CHK_COOL_METHOD_WATER.ForeColor = &H808080
   CHK_COOL_METHOD_WATER.Value = ssCBUnchecked
   
End Sub

Private Sub Chk_Cool_Method_Water_Click()
   
   If CHK_COOL_METHOD_WATER.Value = ssCBUnchecked Then
       If CHK_COOL_METHOD_AIR.Value = ssCBUnchecked Then
         ' CHK_COOL_METHOD_WATER.Value = ssCBChecked
          TXT_COOLING_METHOD.Text = ""
          CHK_COOL_METHOD_WATER.ForeColor = &H80000012
          CHK_COOL_METHOD_AIR.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_COOLING_METHOD.Text = "W"
   
   CHK_COOL_METHOD_WATER.ForeColor = &HFF&
   CHK_COOL_METHOD_WATER.Value = ssCBChecked

   CHK_COOL_METHOD_AIR.ForeColor = &H808080
   CHK_COOL_METHOD_AIR.Value = ssCBUnchecked
   
End Sub

Private Sub Chk_Cool_Prty_Coil_Click()
  
   If CHK_COOL_PRTY_COIL.Value = ssCBUnchecked Then
       If CHK_COOL_PRTY_FIN.Value = ssCBUnchecked Then
        '  CHK_COOL_PRTY_COIL.Value = ssCBChecked
          TXT_COOL_PRTY.Text = ""
          CHK_COOL_PRTY_COIL.ForeColor = &H80000012
          CHK_COOL_PRTY_FIN.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_COOL_PRTY.Text = "C"
   
   CHK_COOL_PRTY_COIL.ForeColor = &HFF&
   CHK_COOL_PRTY_COIL.Value = ssCBChecked

   CHK_COOL_PRTY_FIN.ForeColor = &H808080
   CHK_COOL_PRTY_FIN.Value = ssCBUnchecked
   
End Sub

Private Sub Chk_Cool_Prty_Fin_Click()
    
   If CHK_COOL_PRTY_FIN.Value = ssCBUnchecked Then
       If CHK_COOL_PRTY_COIL.Value = ssCBUnchecked Then
       '   CHK_COOL_PRTY_FIN.Value = ssCBChecked
          TXT_COOL_PRTY.Text = ""
          CHK_COOL_PRTY_FIN.ForeColor = &H80000012
          CHK_COOL_PRTY_COIL.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_COOL_PRTY.Text = "F"
   
   CHK_COOL_PRTY_FIN.ForeColor = &HFF&
   CHK_COOL_PRTY_FIN.Value = ssCBChecked

   CHK_COOL_PRTY_COIL.ForeColor = &H808080
   CHK_COOL_PRTY_COIL.Value = ssCBUnchecked
  
End Sub

Private Sub Chk_Cr_Cd_Click()
   
   If CHK_CR_CD.Value = ssCBUnchecked Then
       If CHK_NON_CR_CD.Value = ssCBUnchecked Then
'          CHK_CR_CD.Value = ssCBChecked
          txt_CR_CD.Text = ""
          CHK_CR_CD.ForeColor = &H80000012
          CHK_NON_CR_CD.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   txt_CR_CD.Text = "1"
   
   CHK_CR_CD.ForeColor = &HFF&
   CHK_CR_CD.Value = ssCBChecked

   CHK_NON_CR_CD.ForeColor = &H808080
   CHK_NON_CR_CD.Value = ssCBUnchecked
      
End Sub

Private Sub Chk_Non_Cr_Cd_Click()
  
   If CHK_NON_CR_CD.Value = ssCBUnchecked Then
       If CHK_CR_CD.Value = ssCBUnchecked Then
'          CHK_NON_CR_CD.Value = ssCBChecked
          txt_CR_CD.Text = ""
          CHK_NON_CR_CD.ForeColor = &H80000012
          CHK_CR_CD.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   txt_CR_CD.Text = "0"
   
   CHK_NON_CR_CD.ForeColor = &HFF&
   CHK_NON_CR_CD.Value = ssCBChecked

   CHK_CR_CD.ForeColor = &H808080
   CHK_CR_CD.Value = ssCBUnchecked
   
End Sub

Private Sub Chk_Prod_Coil_Click()
  
   If CHK_PROD_COIL.Value = ssCBUnchecked Then
       If CHK_PROD_PLATE.Value = ssCBUnchecked And CHK_PROD_C_PLATE.Value = ssCBUnchecked And CHK_PROD_C_COIL.Value = ssCBUnchecked Then
 '         CHK_PROD_COIL.Value = ssCBChecked
          TXT_PROD_STATUS.Text = ""
          CHK_PROD_COIL.ForeColor = &H80000012
          CHK_PROD_PLATE.ForeColor = &H80000012
          CHK_PROD_C_PLATE.ForeColor = &H80000012
          CHK_PROD_C_COIL.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_PROD_STATUS.Text = "2"
   
   CHK_PROD_COIL.ForeColor = &HFF&
   CHK_PROD_COIL.Value = ssCBChecked

   CHK_PROD_PLATE.ForeColor = &H808080
   CHK_PROD_PLATE.Value = ssCBUnchecked

   CHK_PROD_C_PLATE.ForeColor = &H808080
   CHK_PROD_C_PLATE.Value = ssCBUnchecked
   
   CHK_PROD_C_COIL.ForeColor = &H808080
   CHK_PROD_C_COIL.Value = ssCBUnchecked

End Sub
Private Sub Chk_Prod_C_Coil_Click()
  
   If CHK_PROD_C_COIL.Value = ssCBUnchecked Then
       If CHK_PROD_PLATE.Value = ssCBUnchecked And CHK_PROD_C_PLATE.Value = ssCBUnchecked And CHK_PROD_COIL.Value = ssCBUnchecked Then
 '         CHK_PROD_COIL.Value = ssCBChecked
          TXT_PROD_STATUS.Text = ""
          CHK_PROD_C_COIL.ForeColor = &H80000012
          CHK_PROD_PLATE.ForeColor = &H80000012
          CHK_PROD_C_PLATE.ForeColor = &H80000012
          CHK_PROD_COIL.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_PROD_STATUS.Text = "1"
   
   CHK_PROD_C_COIL.ForeColor = &HFF&
   CHK_PROD_C_COIL.Value = ssCBChecked

   CHK_PROD_PLATE.ForeColor = &H808080
   CHK_PROD_PLATE.Value = ssCBUnchecked

   CHK_PROD_C_PLATE.ForeColor = &H808080
   CHK_PROD_C_PLATE.Value = ssCBUnchecked
   
   CHK_PROD_COIL.ForeColor = &H808080
   CHK_PROD_COIL.Value = ssCBUnchecked

End Sub

Private Sub Chk_Prod_Plate_Click()
   
   If CHK_PROD_PLATE.Value = ssCBUnchecked Then
       If CHK_PROD_COIL.Value = ssCBUnchecked And CHK_PROD_C_PLATE.Value = ssCBUnchecked And CHK_PROD_C_COIL.Value = ssCBUnchecked Then
       '   CHK_PROD_PLATE.Value = ssCBChecked
          TXT_PROD_STATUS.Text = ""
          CHK_PROD_PLATE.ForeColor = &H80000012
          CHK_PROD_COIL.ForeColor = &H80000012
          CHK_PROD_C_COIL.ForeColor = &H80000012
          CHK_PROD_C_PLATE.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_PROD_STATUS.Text = "0"
   
   CHK_PROD_PLATE.ForeColor = &HFF&
   CHK_PROD_PLATE.Value = ssCBChecked

   CHK_PROD_COIL.ForeColor = &H808080
   CHK_PROD_COIL.Value = ssCBUnchecked
   
   CHK_PROD_C_COIL.ForeColor = &H808080
   CHK_PROD_C_COIL.Value = ssCBUnchecked
   
   CHK_PROD_C_PLATE.ForeColor = &H808080
   CHK_PROD_C_PLATE.Value = ssCBUnchecked
        
End Sub
Private Sub Chk_Prod_C_Plate_Click()
   
   If CHK_PROD_C_PLATE.Value = ssCBUnchecked Then
       If CHK_PROD_COIL.Value = ssCBUnchecked And CHK_PROD_PLATE.Value = ssCBUnchecked And CHK_PROD_C_COIL.Value = ssCBUnchecked Then
       '   CHK_PROD_PLATE.Value = ssCBChecked
          TXT_PROD_STATUS.Text = ""
          CHK_PROD_C_PLATE.ForeColor = &H80000012
          CHK_PROD_COIL.ForeColor = &H80000012
          CHK_PROD_C_COIL.ForeColor = &H80000012
          CHK_PROD_PLATE.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_PROD_STATUS.Text = "3"
   
   CHK_PROD_C_PLATE.ForeColor = &HFF&
   CHK_PROD_C_PLATE.Value = ssCBChecked

   CHK_PROD_COIL.ForeColor = &H808080
   CHK_PROD_COIL.Value = ssCBUnchecked
   
   CHK_PROD_C_COIL.ForeColor = &H808080
   CHK_PROD_C_COIL.Value = ssCBUnchecked
   
   CHK_PROD_PLATE.ForeColor = &H808080
   CHK_PROD_PLATE.Value = ssCBUnchecked
        
End Sub

Private Sub Chk_Rolling_Auto_Click()
   
  If CHK_ROLLING_AUTO.Value = ssCBUnchecked Then
       If CHK_ROLLING_OP.Value = ssCBUnchecked Then
         ' CHK_ROLLING_AUTO.Value = ssCBChecked
          TXT_ROLLING_METHOD.Text = ""
          CHK_ROLLING_AUTO.ForeColor = &H80000012
          CHK_ROLLING_OP.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_ROLLING_METHOD.Text = "0"
   
   CHK_ROLLING_AUTO.ForeColor = &HFF&
   CHK_ROLLING_AUTO.Value = ssCBChecked

   CHK_ROLLING_OP.ForeColor = &H808080
   CHK_ROLLING_OP.Value = ssCBUnchecked
  
End Sub

Private Sub Chk_Rolling_Op_Click()
  
  If CHK_ROLLING_OP.Value = ssCBUnchecked Then
       If CHK_ROLLING_AUTO.Value = ssCBUnchecked Then
         ' CHK_ROLLING_OP.Value = ssCBChecked
          TXT_ROLLING_METHOD.Text = ""
          CHK_ROLLING_OP.ForeColor = &H80000012
          CHK_ROLLING_AUTO.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_ROLLING_METHOD.Text = "1"
   
   CHK_ROLLING_OP.ForeColor = &HFF&
   CHK_ROLLING_OP.Value = ssCBChecked

   CHK_ROLLING_AUTO.ForeColor = &H808080
   CHK_ROLLING_AUTO.Value = ssCBUnchecked
      
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

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gf_Mill_ComboAdd(M_CN1, CBO_SLAB_NO, "CB")
    
    If CBO_SLAB_NO.ListCount <> 0 Then
       CBO_SLAB_NO.ListIndex = 0
    End If
    
    TXT_UPD_EMP = sUserID '+ ":" + sUsername
    CBO_PLT.ListIndex = 0
     
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
    
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
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set sControl = Nothing
    Set MC = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

   Select Case SSTab1.Tab
          Case 0
              Call Gp_Ms_Cls(Mc1("rControl"))
              Call Gp_SSCheck_Cls(MC("sControl"))
          Case 1
              Call Gp_Ms_Cls(Mc2("rControl"))
          
   End Select

    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)

    Mc1("pControl")(1).SetFocus
    CBO_PLT.ListIndex = 0
    TXT_UPD_EMP = sUserID
    
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
    
    Call Gf_Mill_ComboAdd(M_CN1, CBO_SLAB_NO, "CB")
    
End Sub

Public Sub Master_Cpy()

    Select Case SSTab1.Tab
          Case 0
            Call Gf_Ms_Copy(Mc1)
          Case 1
            Call Gf_Ms_Copy(Mc2)
   End Select
   
End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If
     
     Select Case SSTab1.Tab
          Case 0
          
                If Gf_Ms_Paste(M_CN1, Mc1) Then
                   Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
                   Call Gp_Ms_ControlLock(Mc1("pControl"), False)
                End If
          Case 1
          
                If Gf_Ms_Paste(M_CN1, Mc2) Then
                   Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
                   Call Gp_Ms_ControlLock(Mc1("pControl"), False)
                End If
   End Select

End Sub

Public Sub Form_Ref()

    Dim S_DAY As String
    Dim S_HOUR As String
    Dim S_MIN As String
    Dim E_DAY As String
    Dim E_HOUR As String
    Dim E_MIN As String
    Dim C1_HOUR As Integer
    Dim C1_MIN As Integer
    Dim C2_HOUR As Integer
    Dim C2_MIN As Integer
    Dim CAL_MIN As Integer
     
    Call Gp_SSCheck_Cls(MC("sControl"))
    
    Select Case SSTab1.Tab
           
           Case 0
     
            If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl")) Then
               Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
               Call Gp_Ms_ControlLock(Mc1("pControl"), True)
               CBO_SLAB_NO.Enabled = True
'               Call Gp_Ms_Cls(Mc2("rControl"))
            End If
            
           Case 1
           
            If Gf_Ms_Refer(M_CN1, Mc2, , , False) Then
               Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
               Call Gp_Ms_ControlLock(Mc2("pControl"), True)
               CBO_SLAB_NO.Enabled = True
'               Call Gp_Ms_Cls(Mc2("rControl"))
            End If
    
    End Select
     
    TXT_UPD_EMP = sUserID '+ ":" + sUsername
    CBO_PLT.ListIndex = 0

    
End Sub

Public Sub Form_Pro()

    Dim SMESG As String
    Dim Temp_no As String
    
    Temp_no = CBO_SLAB_NO.Text
    
    TXT_UPD_EMP = sUserID

    Select Case SSTab1.Tab
    
          Case 0
          
                 If Not Gp_DateCheck(TXT_MILL_STA_TIME) Then
                      SMESG = " 请正确输入开轧时间 ！"
                      Call Gp_MsgBoxDisplay(SMESG)
                      Exit Sub
                 End If
                 
                 If TXT_MILL_STA_TIME.RawData = "" And TXT_MILL_END_TIME.RawData = "" Then
                      SMESG = " 请输入开轧时间 ！"
                      Call Gp_MsgBoxDisplay(SMESG)
                      Exit Sub
                 ElseIf TXT_MILL_STA_TIME.RawData = "" And TXT_MILL_END_TIME.RawData <> "" Then
                      SMESG = " 请首先输入开轧时间 ！"
                      Call Gp_MsgBoxDisplay(SMESG)
                      Exit Sub
                 ElseIf TXT_MILL_STA_TIME.RawData <> "" And TXT_MILL_END_TIME.RawData <> "" Then
                        If Not Gp_DateCheck(TXT_MILL_END_TIME) Then
                             SMESG = " 请正确输入终轧时间 ！"
                             Call Gp_MsgBoxDisplay(SMESG)
                             Exit Sub
                        End If
                        If Val(TXT_MILL_STA_TIME.RawData) - Val(TXT_MILL_END_TIME.RawData) > 0 Then
                             SMESG = " 终轧时间应大于开轧时间，请正确输入时间信息 ！"
                             Call Gp_MsgBoxDisplay(SMESG)
                             Exit Sub
                        End If
                 End If
                      
                 If Trim(txt_CR_CD) = "1" Then
                    If Trim(SDB_CR_STAGE1_THK) = "" And Trim(SDB_CR_STAGE1_TEMP) = "" And Trim(SDB_CR_STAGE1_TIME) = "" Then
                        SMESG = " 请输入控轧一阶段厚度，温度，待轧时间 ！"
                        Call Gp_MsgBoxDisplay(SMESG)
                        Exit Sub
                    End If
                 Else
                    SDB_CR_STAGE1_THK = ""
                    SDB_CR_STAGE1_TEMP = ""
                    SDB_CR_STAGE1_TIME = ""
                    SDB_CR_STAGE2_THK = ""
                    SDB_CR_STAGE2_TEMP = ""
                    SDB_CR_STAGE2_TIME = ""
                    SDB_CR_STAGE3_THK = ""
                    SDB_CR_STAGE3_TEMP = ""
                    SDB_CR_STAGE3_TIME = ""
                 End If
                 If Gf_Mc_Authority(sAuthority, Mc1) Then
                   ' txt_ins_emp.Text = sUserID
                   If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                      CBO_SLAB_NO.Enabled = True
                 End If
                 CBO_SLAB_NO.Text = Temp_no
                 
                SMESG = "请确认母板实绩是否录入"
                Call Gp_MsgBoxDisplay(SMESG, "I", "重要提示")
                 
          Case 1
                 If TXT_CUTEND_CD.Text = "Y" And TXT_COMFRM.Text = "2" Then
                    SMESG = " （缺号母板确定） 与 （母板剪切结束确定）不能同时操作 ！"
                    Call Gp_MsgBoxDisplay(SMESG)
                    Exit Sub
                 End If
                 
                 If TXT_CUTEND_CD.Text = "Y" Then
                    SMESG = " 确定此轧件剪切母板结束 ？ "
                 ElseIf TXT_COMFRM.Text = "2" Then
                    SMESG = " 确定以下母板缺号 ？ "
                 Else
                    If Gf_Mc_Authority(sAuthority, Mc1) Then
                       If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                       CBO_SLAB_NO.Enabled = True
                    End If
                    Exit Sub
                 End If
                 If Gp_MsgBox(SMESG, "C") = 6 Then
                    If Gf_Mc_Authority(sAuthority, Mc1) Then
                       If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                       CBO_SLAB_NO.Enabled = True
                    End If
                 End If
      
   End Select
   
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

Private Sub SDB_MOTHER_PLATE_LEN1_DblClick()
    If SDB_MOTHER_PLATE_LEN1.Text = "" Then
        SDB_MOTHER_PLATE_LEN1.Text = SDB_MOTHER_SCH_LEN1.Text
    Else
        SDB_MOTHER_PLATE_LEN1.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN10_DblClick()
    If SDB_MOTHER_PLATE_LEN10.Text = "" Then
        SDB_MOTHER_PLATE_LEN10.Text = SDB_MOTHER_SCH_LEN10.Text
    Else
        SDB_MOTHER_PLATE_LEN10.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN11_DblClick()
    If SDB_MOTHER_PLATE_LEN11.Text = "" Then
        SDB_MOTHER_PLATE_LEN11.Text = SDB_MOTHER_SCH_LEN11.Text
    Else
        SDB_MOTHER_PLATE_LEN11.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN12_DblClick()
    If SDB_MOTHER_PLATE_LEN12.Text = "" Then
        SDB_MOTHER_PLATE_LEN12.Text = SDB_MOTHER_SCH_LEN12.Text
    Else
        SDB_MOTHER_PLATE_LEN12.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN2_DblClick()
    If SDB_MOTHER_PLATE_LEN2.Text = "" Then
        SDB_MOTHER_PLATE_LEN2.Text = SDB_MOTHER_SCH_LEN2.Text
    Else
        SDB_MOTHER_PLATE_LEN2.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN3_DblClick()
    If SDB_MOTHER_PLATE_LEN3.Text = "" Then
        SDB_MOTHER_PLATE_LEN3.Text = SDB_MOTHER_SCH_LEN3.Text
    Else
        SDB_MOTHER_PLATE_LEN3.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN4_DblClick()
    If SDB_MOTHER_PLATE_LEN4.Text = "" Then
        SDB_MOTHER_PLATE_LEN4.Text = SDB_MOTHER_SCH_LEN4.Text
    Else
        SDB_MOTHER_PLATE_LEN4.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN5_DblClick()
    If SDB_MOTHER_PLATE_LEN5.Text = "" Then
        SDB_MOTHER_PLATE_LEN5.Text = SDB_MOTHER_SCH_LEN5.Text
    Else
        SDB_MOTHER_PLATE_LEN5.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN6_DblClick()
    If SDB_MOTHER_PLATE_LEN6.Text = "" Then
        SDB_MOTHER_PLATE_LEN6.Text = SDB_MOTHER_SCH_LEN6.Text
    Else
        SDB_MOTHER_PLATE_LEN6.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN7_DblClick()
    If SDB_MOTHER_PLATE_LEN7.Text = "" Then
        SDB_MOTHER_PLATE_LEN7.Text = SDB_MOTHER_SCH_LEN7.Text
    Else
        SDB_MOTHER_PLATE_LEN7.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN8_DblClick()
    If SDB_MOTHER_PLATE_LEN8.Text = "" Then
        SDB_MOTHER_PLATE_LEN8.Text = SDB_MOTHER_SCH_LEN8.Text
    Else
        SDB_MOTHER_PLATE_LEN8.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN9_DblClick()
    If SDB_MOTHER_PLATE_LEN9.Text = "" Then
        SDB_MOTHER_PLATE_LEN9.Text = SDB_MOTHER_SCH_LEN9.Text
    Else
        SDB_MOTHER_PLATE_LEN9.Text = ""
    End If
End Sub

Private Sub SSCommand1_Click()
    If Trim(TXT_COMFRM.Text) = "" Then
       TXT_COMFRM.Text = "2"
    Else
       TXT_COMFRM.Text = ""
    End If
End Sub

Private Sub SSCommand2_Click()
    If Trim(TXT_CUTEND_CD.Text) = "" Then
       TXT_CUTEND_CD.Text = "Y"
    Else
       TXT_CUTEND_CD.Text = ""
    End If
End Sub

Private Sub TXT_MILL_STA_TIME_DblClick()

    TXT_MILL_STA_TIME.RawData = Gf_DTSet(M_CN1) 'Format(Now, "YYYYMMDDHHMMSS")

End Sub

Private Sub TXT_MILL_END_TIME_DblClick()

    TXT_MILL_END_TIME.RawData = Gf_DTSet(M_CN1) 'Format(Now, "YYYYMMDDHHMMSS")

End Sub
Private Sub CHK_YES_Click()

    If CHK_YES.Value = ssCBUnchecked Then
        If CHK_NO.Value = ssCBUnchecked Then
          ' CHK_YES.Value = ssCBChecked
           TXT_COMFRM.Text = ""
           CHK_YES.ForeColor = &H80000012
           CHK_NO.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_COMFRM.Text = "1"
    
    CHK_YES.ForeColor = &HFF&
    CHK_YES.Value = ssCBChecked
    
    CHK_NO.ForeColor = &H808080
    CHK_NO.Value = ssCBUnchecked
   
End Sub
Private Sub CHK_NO_Click()

    If CHK_NO.Value = ssCBUnchecked Then
        If CHK_YES.Value = ssCBUnchecked Then
          ' CHK_NO.Value = ssCBChecked
           TXT_COMFRM.Text = ""
           CHK_NO.ForeColor = &H80000012
           CHK_YES.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_COMFRM.Text = "2"
    
    CHK_NO.ForeColor = &HFF&
    CHK_NO.Value = ssCBChecked
    
    CHK_YES.ForeColor = &H808080
    CHK_YES.Value = ssCBUnchecked
   
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Mill_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant,sPRC String,
'                    {sFACT_CD,sPRC_LINE String, sADDNUM As Integer, ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : combo Add
'---------------------------------------------------------------------------------------
Private Function Gf_Roll_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPrc As String, Optional sFACT_CD As String = "C1", _
             Optional sPRC_LINE As String = "1", Optional sADDNUM As Integer = 5, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String

    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Roll_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT GOODS_ID FROM (SELECT B.GOODS_ID "
    sQuery = sQuery + "               FROM FP_TRACKIDX A, FP_TRACKDATA B "
    sQuery = sQuery + "              WHERE A.FACT_CD  = '" + sFACT_CD + "'"
    sQuery = sQuery + "                AND A.PRC      = '" + sPrc + "'"
    sQuery = sQuery + "                AND A.PRC_LINE = '" + sPRC_LINE + "'"
    sQuery = sQuery + "                AND A.FACT_CD  =B.FACT_CD "
    sQuery = sQuery + "                AND A.PRC      =B.PRC "
    sQuery = sQuery + "                AND A.PRC_LINE =B.PRC_LINE "
    sQuery = sQuery + "                AND B.SEQ_NO  <= A.LAST_SEQ "
    sQuery = sQuery + "           ORDER BY B.SEQ_NO DESC) "
    sQuery = sQuery + "              WHERE ROWNUM    <= " + CStr(sADDNUM)

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_Roll_ComboAdd = True
    Else
        Gf_Roll_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_Roll_ComboAdd = False

End Function



