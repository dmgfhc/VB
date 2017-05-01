VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0032C 
   Caption         =   "产品检验实绩录入（金相）_AQC0032C"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8730
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   15399
      _Version        =   196609
      AutoSize        =   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "AQC0032C.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   7365
         Left            =   0
         TabIndex        =   3
         Top             =   1365
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   12991
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabHeight       =   520
         TabCaption(0)   =   "金相试验录入"
         TabPicture(0)   =   "AQC0032C.frx":0072
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSPanel3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSPanel4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SSPanel5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         Begin Threed.SSPanel SSPanel5 
            Height          =   7365
            Left            =   10080
            TabIndex        =   33
            Top             =   480
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   12991
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_TIN_GRD 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   85
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   4650
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.TextBox txt_DS_GRD 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   84
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   4245
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_BRST4 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   83
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   3810
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_BRST3 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   82
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   3390
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_BRST2 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   81
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   2985
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_BRST1 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   80
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   2580
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_ARST4 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   79
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   1980
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_ARST3 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   78
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   1620
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_ARST2 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   77
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   1215
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_ARST1 
               Height          =   315
               Left            =   3780
               MaxLength       =   80
               TabIndex        =   76
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   750
               Width           =   900
            End
            Begin VB.TextBox txt_NON_METAL_ACD3_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   41
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   1620
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_ACD2_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   40
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   1215
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_ACD1_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   39
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   750
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_ACD3 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   20
               Tag             =   "8"
               Top             =   1620
               Width           =   435
            End
            Begin VB.TextBox txt_NON_METAL_ACD2 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   19
               Tag             =   "8"
               Top             =   1215
               Width           =   435
            End
            Begin VB.TextBox txt_NON_METAL_ACD1 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   18
               Tag             =   "8"
               Top             =   750
               Width           =   435
            End
            Begin VB.TextBox txt_NON_METAL_ACD4_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   38
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   1980
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_ACD4 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   21
               Tag             =   "8"
               Top             =   1980
               Width           =   435
            End
            Begin VB.TextBox txt_NON_METAL_BCD3_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   37
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   3390
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_BCD2_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   36
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   2985
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_BCD1_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   35
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   2580
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_BCD3 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   24
               Tag             =   "8"
               Top             =   3390
               Width           =   435
            End
            Begin VB.TextBox txt_NON_METAL_BCD2 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   23
               Tag             =   "8"
               Top             =   2985
               Width           =   435
            End
            Begin VB.TextBox txt_NON_METAL_BCD1 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   22
               Tag             =   "8"
               Top             =   2580
               Width           =   435
            End
            Begin VB.TextBox txt_NON_METAL_BCD4_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               MaxLength       =   80
               TabIndex        =   34
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   3810
               Width           =   1605
            End
            Begin VB.TextBox txt_NON_METAL_BCD4 
               Height          =   315
               Left            =   1650
               MaxLength       =   1
               TabIndex        =   25
               Tag             =   "8"
               Top             =   3810
               Width           =   435
            End
            Begin InDate.ULabel ULabel87 
               Height          =   1770
               Index           =   1
               Left            =   1020
               Top             =   660
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
               Left            =   1020
               Top             =   2490
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
               Height          =   3915
               Index           =   25
               Left            =   180
               Top             =   660
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   6906
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
               Left            =   120
               Top             =   210
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
               Left            =   3780
               Top             =   210
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
               Left            =   1605
               Top             =   210
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
            Begin InDate.ULabel ULabel87 
               Height          =   330
               Index           =   0
               Left            =   1020
               Top             =   4230
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
            Begin InDate.ULabel ULabel87 
               Height          =   330
               Index           =   3
               Left            =   1020
               Top             =   4635
               Visible         =   0   'False
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
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   7365
            Left            =   5160
            TabIndex        =   42
            Top             =   480
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   12991
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_JOMINY_RST_TOP3 
               Height          =   315
               Left            =   3840
               TabIndex        =   75
               Tag             =   "7"
               Top             =   4725
               Width           =   900
            End
            Begin VB.TextBox txt_JOMINY_RST_TOP2 
               Height          =   315
               Left            =   3840
               TabIndex        =   74
               Tag             =   "7"
               Top             =   4320
               Width           =   900
            End
            Begin VB.TextBox txt_JOMINY_RST_TOP1 
               Height          =   315
               Left            =   3840
               TabIndex        =   73
               Tag             =   "7"
               Top             =   3975
               Width           =   900
            End
            Begin VB.TextBox txt_BELT_STR_GRD_RST 
               Height          =   315
               Left            =   3840
               TabIndex        =   72
               Tag             =   "6"
               Top             =   3060
               Width           =   900
            End
            Begin VB.TextBox txt_ACD_RST5 
               Height          =   315
               Left            =   3840
               TabIndex        =   71
               Tag             =   "5"
               Top             =   2490
               Width           =   900
            End
            Begin VB.TextBox txt_ACD_RST4 
               Height          =   315
               Left            =   3840
               TabIndex        =   70
               Tag             =   "5"
               Top             =   2055
               Width           =   900
            End
            Begin VB.TextBox txt_ACD_RST3 
               Height          =   315
               Left            =   3840
               TabIndex        =   69
               Tag             =   "5"
               Top             =   1665
               Width           =   900
            End
            Begin VB.TextBox txt_ACD_RST2 
               Height          =   315
               Left            =   3840
               TabIndex        =   68
               Tag             =   "5"
               Top             =   1260
               Width           =   900
            End
            Begin VB.TextBox txt_ACD_RST1 
               Height          =   315
               Left            =   3840
               TabIndex        =   67
               Tag             =   "5"
               Top             =   810
               Width           =   900
            End
            Begin VB.TextBox txt_JOMINY_NAME 
               Height          =   300
               Left            =   1680
               TabIndex        =   17
               Tag             =   "7"
               Top             =   4350
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.TextBox txt_JOMINY_TYP 
               Height          =   300
               Left            =   1680
               TabIndex        =   16
               Tag             =   "7"
               Top             =   3975
               Width           =   735
            End
            Begin VB.TextBox txt_ACD_DFT_TYP1 
               Height          =   315
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   11
               Tag             =   "5"
               Top             =   810
               Width           =   435
            End
            Begin VB.TextBox txt_ACD_DFT_TYP2 
               Height          =   315
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   12
               Tag             =   "5"
               Top             =   1245
               Width           =   435
            End
            Begin VB.TextBox txt_ACD_DFT_TYP3 
               Height          =   315
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   13
               Tag             =   "5"
               Top             =   1650
               Width           =   435
            End
            Begin VB.TextBox txt_ACD_DFT_TYP1_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2130
               MaxLength       =   80
               TabIndex        =   47
               TabStop         =   0   'False
               Tag             =   "5"
               Top             =   810
               Width           =   1605
            End
            Begin VB.TextBox txt_ACD_DFT_TYP2_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2130
               MaxLength       =   80
               TabIndex        =   46
               TabStop         =   0   'False
               Tag             =   "5"
               Top             =   1245
               Width           =   1605
            End
            Begin VB.TextBox txt_ACD_DFT_TYP3_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2130
               MaxLength       =   80
               TabIndex        =   45
               TabStop         =   0   'False
               Tag             =   "5"
               Top             =   1650
               Width           =   1605
            End
            Begin VB.TextBox txt_ACD_DFT_TYP5_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2130
               MaxLength       =   80
               TabIndex        =   44
               TabStop         =   0   'False
               Tag             =   "5"
               Top             =   2475
               Width           =   1605
            End
            Begin VB.TextBox txt_ACD_DFT_TYP4_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2130
               MaxLength       =   80
               TabIndex        =   43
               TabStop         =   0   'False
               Tag             =   "5"
               Top             =   2040
               Width           =   1605
            End
            Begin VB.TextBox txt_ACD_DFT_TYP5 
               Height          =   315
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   15
               Tag             =   "5"
               Top             =   2475
               Width           =   435
            End
            Begin VB.TextBox txt_ACD_DFT_TYP4 
               Height          =   315
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   14
               Tag             =   "5"
               Top             =   2040
               Width           =   435
            End
            Begin InDate.ULabel ULabel27 
               Height          =   315
               Index           =   23
               Left            =   180
               Top             =   810
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
               Left            =   1665
               Top             =   210
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
               Left            =   180
               Top             =   210
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
               Left            =   3840
               Top             =   210
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
               Left            =   180
               Top             =   3060
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
            Begin CSTextLibCtl.sidbEdit sdb_JOMINY_DIST1 
               Height          =   315
               Left            =   2430
               TabIndex        =   48
               Tag             =   "7"
               Top             =   3960
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
               Left            =   2430
               TabIndex        =   49
               Tag             =   "7"
               Top             =   4335
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
               Left            =   2430
               TabIndex        =   50
               Tag             =   "7"
               Top             =   4710
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
               Left            =   180
               Top             =   3975
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
               Left            =   180
               Top             =   3600
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
               Left            =   3840
               Top             =   3600
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
               Left            =   1635
               Top             =   3600
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
               Left            =   2430
               Top             =   3600
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
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   7365
            Left            =   0
            TabIndex        =   51
            Top             =   480
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   12991
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_FRACT_GRD_RST3 
               Height          =   315
               Left            =   3780
               TabIndex        =   66
               Tag             =   "4"
               Top             =   3765
               Width           =   930
            End
            Begin VB.TextBox txt_FRACT_GRD_RST2 
               Height          =   315
               Left            =   3780
               TabIndex        =   65
               Tag             =   "4"
               Top             =   3360
               Width           =   930
            End
            Begin VB.TextBox txt_FRACT_GRD_RST5 
               Height          =   315
               Left            =   3780
               TabIndex        =   64
               Tag             =   "4"
               Top             =   4590
               Width           =   930
            End
            Begin VB.TextBox txt_FRACT_GRD_RST4 
               Height          =   315
               Left            =   3780
               TabIndex        =   63
               Tag             =   "4"
               Top             =   4185
               Width           =   930
            End
            Begin VB.TextBox txt_FRACT_GRD_RST1 
               Height          =   315
               Left            =   3780
               TabIndex        =   62
               Tag             =   "4"
               Top             =   2955
               Width           =   930
            End
            Begin VB.TextBox txt_S_PRINT_RST 
               Height          =   315
               Left            =   3780
               TabIndex        =   61
               Tag             =   "3"
               Top             =   2325
               Width           =   930
            End
            Begin VB.TextBox txt_RMV_CAR_RST 
               Height          =   315
               Left            =   3780
               TabIndex        =   60
               Tag             =   "2"
               Top             =   1725
               Width           =   930
            End
            Begin VB.TextBox txt_OST_GRAIN_SIZE_RST 
               Height          =   315
               Left            =   3780
               TabIndex        =   59
               Tag             =   "9"
               Top             =   1260
               Width           =   930
            End
            Begin VB.TextBox txt_GRAIN_SIZE_RST 
               Height          =   315
               Left            =   3780
               TabIndex        =   58
               Tag             =   "1"
               Top             =   840
               Width           =   930
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
               Left            =   1620
               MaxLength       =   1
               TabIndex        =   5
               Tag             =   "2"
               Top             =   1740
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
               Left            =   2130
               MaxLength       =   80
               TabIndex        =   57
               TabStop         =   0   'False
               Tag             =   "2"
               Top             =   1740
               Width           =   1470
            End
            Begin VB.TextBox txt_FRACT_NAME_CD1 
               Height          =   315
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   6
               Tag             =   "4"
               Top             =   2955
               Width           =   435
            End
            Begin VB.TextBox txt_FRACT_NAME_CD2 
               Height          =   315
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   7
               Tag             =   "4"
               Top             =   3360
               Width           =   435
            End
            Begin VB.TextBox txt_FRACT_NAME_CD3 
               Height          =   315
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   8
               Tag             =   "4"
               Top             =   3765
               Width           =   435
            End
            Begin VB.TextBox txt_FRACT_NAME_CD1_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2070
               MaxLength       =   80
               TabIndex        =   56
               TabStop         =   0   'False
               Tag             =   "4"
               Top             =   2955
               Width           =   1605
            End
            Begin VB.TextBox txt_FRACT_NAME_CD2_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2070
               MaxLength       =   80
               TabIndex        =   55
               TabStop         =   0   'False
               Tag             =   "4"
               Top             =   3360
               Width           =   1605
            End
            Begin VB.TextBox txt_FRACT_NAME_CD3_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2070
               MaxLength       =   80
               TabIndex        =   54
               TabStop         =   0   'False
               Tag             =   "4"
               Top             =   3765
               Width           =   1605
            End
            Begin VB.TextBox txt_FRACT_NAME_CD5_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2070
               MaxLength       =   80
               TabIndex        =   53
               TabStop         =   0   'False
               Tag             =   "4"
               Top             =   4590
               Width           =   1605
            End
            Begin VB.TextBox txt_FRACT_NAME_CD4_NAME 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2070
               MaxLength       =   80
               TabIndex        =   52
               TabStop         =   0   'False
               Tag             =   "4"
               Top             =   4185
               Width           =   1605
            End
            Begin VB.TextBox txt_FRACT_NAME_CD5 
               Height          =   315
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   10
               Tag             =   "4"
               Top             =   4590
               Width           =   435
            End
            Begin VB.TextBox txt_FRACT_NAME_CD4 
               Height          =   315
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   9
               Tag             =   "4"
               Top             =   4185
               Width           =   435
            End
            Begin InDate.ULabel ULabel4 
               Height          =   300
               Index           =   0
               Left            =   120
               Top             =   210
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
               Left            =   120
               Top             =   810
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
               Left            =   3780
               Top             =   210
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
               Left            =   1605
               Top             =   210
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
               Left            =   120
               Top             =   1725
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
               Left            =   120
               Top             =   2325
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
               Left            =   120
               Top             =   2955
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
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   8
               Left            =   120
               Top             =   1260
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
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   1005
         _Version        =   196609
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton Op_CHAGE 
            Caption         =   "按炉号保存"
            Height          =   315
            Left            =   6390
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   120
            Width           =   1245
         End
         Begin VB.OptionButton Op_ONLY 
            Caption         =   "单独保存"
            Height          =   315
            Left            =   7800
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   120
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.TextBox txt_SAVE_CASE 
            Height          =   270
            Left            =   9390
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "99"
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   315
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
            Left            =   1500
            MaxLength       =   14
            TabIndex        =   28
            Tag             =   "1"
            Top             =   120
            Width           =   2655
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
            Left            =   5670
            MaxLength       =   1
            TabIndex        =   2
            Tag             =   "取样位置"
            Top             =   120
            Width           =   435
         End
         Begin VB.TextBox txt_smp_loc_p 
            Height          =   345
            Left            =   8220
            TabIndex        =   27
            TabStop         =   0   'False
            Tag             =   "取样位置"
            Top             =   150
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txt_INS_EMP 
            Height          =   375
            Left            =   6240
            TabIndex        =   26
            TabStop         =   0   'False
            Tag             =   "INS_EMP"
            Top             =   150
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txt_smp_no_p 
            Height          =   315
            Left            =   6930
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "试样号"
            Top             =   0
            Visible         =   0   'False
            Width           =   1245
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   0
            Left            =   120
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
            Left            =   4290
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
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   0
         TabIndex        =   32
         Top             =   645
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   1138
         _Version        =   196609
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   2
            Left            =   120
            Top             =   0
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
            Left            =   2070
            Top             =   0
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
            Left            =   4800
            Top             =   0
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
            Left            =   8070
            Top             =   0
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
            Left            =   10020
            Top             =   0
            Width           =   600
            _ExtentX        =   1058
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
            Left            =   10620
            Top             =   0
            Width           =   1950
            _ExtentX        =   3440
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
            Left            =   12570
            Top             =   0
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
            Left            =   13770
            Top             =   0
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
            Left            =   3510
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "取样日期"
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
            Left            =   120
            Top             =   300
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
            Left            =   2070
            Top             =   300
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
            Left            =   4800
            Top             =   300
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
            Left            =   8070
            Top             =   300
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
            Left            =   10020
            Top             =   300
            Width           =   645
            _ExtentX        =   1138
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
            Left            =   10650
            Top             =   300
            Width           =   1935
            _ExtentX        =   3413
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
            Left            =   12570
            Top             =   300
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
            Left            =   13770
            Top             =   300
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
            Left            =   3510
            Top             =   300
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
            Left            =   6810
            Top             =   0
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
            Left            =   6810
            Top             =   300
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
      End
   End
End
Attribute VB_Name = "AQC0032C"
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
'-- Program Name      材质试验实绩输入-金相组
'-- Program ID        AQC0032C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Li Qing Yu
'-- Coder             Li Qing Yu
'-- Date              2006.12.03
'-- Description       材质试验实绩输入
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'
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

Dim sOldAuthority As String         'Save First Load Authority
Dim bExpo_SMP   As Boolean          'This sampling is Expo sampling when value is true


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'TOP and STAND
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call Gp_Ms_Collection(txt_smp_no_p, "p", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_HEAT_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_Cut_DD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_STD_YY, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_ORD_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_ORD_ITEM, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

    'MASTER2 Collection
     Mc2.Add Item:="AQC0032C.P_REFER_HEAD", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'试样号&取样位置
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
           Call Gp_Ms_Collection(txt_smp_no_p, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_smp_loc_p, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'金相检验 - TAB 3
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call Gp_Ms_Collection(txt_GRAIN_SIZE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_RMV_CAR_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_RMV_CAR_TYP_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_RMV_CAR_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_S_PRINT_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_FRACT_GRD_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_FRACT_GRD_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_FRACT_GRD_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_FRACT_GRD_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_FRACT_NAME_CD5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_FRACT_NAME_CD5_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_FRACT_GRD_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
                    
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ACD_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ACD_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ACD_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ACD_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
           Call Gp_Ms_Collection(txt_ACD_DFT_TYP5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ACD_DFT_TYP5_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ACD_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
       Call Gp_Ms_Collection(txt_BELT_STR_GRD_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
            Call Gp_Ms_Collection(txt_JOMINY_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_JOMINY_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_JOMINY_DIST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_JOMINY_DIST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_JOMINY_DIST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_JOMINY_RST_TOP1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_JOMINY_RST_TOP2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_JOMINY_RST_TOP3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
                
         Call Gp_Ms_Collection(txt_NON_METAL_ACD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_ARST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_ACD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_ARST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_ACD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_ARST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_ACD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_ACD4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_ARST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_BRST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_BRST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_BRST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         Call Gp_Ms_Collection(txt_NON_METAL_BCD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_NON_METAL_BCD4_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NON_METAL_BRST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------- Master End ------------------------------------------------------------------------------------
                Call Gp_Ms_Collection(txt_SAVE_CASE, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'              Call Gp_Ms_Collection(txt_INPUT_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'                Call Gp_Ms_Collection(txt_UPD_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_OST_GRAIN_SIZE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_DS_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_TIN_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
     Mc1.Add Item:="AQC0032C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQC0032C.P_REFER", Key:="P-R"
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

''---------------------------------------------------------------------------------------------------------------------------------------------
''--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
''---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String

    Select Case Me.ActiveControl.Name

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
'

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub
'
Private Sub Form_KeyPress(KeyAscii As Integer)


    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = 19 Or KeyAscii = 10 Then
        KeyAscii = 0
        Call Form_Pro
    End If


End Sub
'
Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    sOldAuthority = sAuthority

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Screen.MousePointer = vbDefault
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
'
Public Sub Form_Exit()

    Unload Me

End Sub
'
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


End Sub
'
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub
'
Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

End Sub
'
Public Sub Form_Ref()
    Dim sMesg           As String
    Dim sSMP_NO         As String
    Dim sPROD_CD        As String
        
        sSMP_NO = Trim(txt_SMP_NO.Text)
        
        sPROD_CD = SMP_PROD_Check(sSMP_NO)
        
        If sPROD_CD = "ER" Then Exit Sub
                
        Call Form_Cls
      
        If Gf_Ms_Refer(M_CN1, Mc2, Mc1("nControl"), Mc1("mControl")) Then
            If sAuthority = "1000" Or sAuthority = "0000" Then
                Call MsgBox("你没有当前试样号：" + sSMP_NO + " 操作权限！", vbOKOnly, "系统提示")
            End If
            Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)

        End If

    Call subItemLock(txt_SMP_NO.Text)

End Sub

Private Sub Op_CHAGE_Click()
    If Op_CHAGE.Value = True Then txt_SAVE_CASE.Text = 1
End Sub

Private Sub Op_ONLY_Click()
    If Op_ONLY.Value = True Then txt_SAVE_CASE.Text = 0
    
End Sub



Private Sub txt_SMP_CUT_LOC_Change()
Dim sPROD_CD As String

    sPROD_CD = SMP_PROD_Check(Trim(txt_smp_no_p.Text))
    
    If sPROD_CD = "ER" Then
        Exit Sub
    Else
        txt_smp_loc_p.Text = Trim(txt_SMP_CUT_LOC.Text)
    End If
End Sub

Public Sub Form_Pro()
  
    
    If Gf_Mc_Authority(sAuthority, Mc1) Then
            txt_ins_emp.Text = sUserID
            If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If

End Sub
'
Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub
'
Private Sub subItemLock(ByVal sSMP_NO As String)
    Dim sQuery          As String
    Dim arrayRecord     As Variant
    Dim AdoRs           As adodb.Recordset

 On Error GoTo Error_Rtn
    Set AdoRs = New adodb.Recordset
    
    sQuery = "{call AQC0032C.P_MART_ITEM_SELECT('" + sSMP_NO + "')}"

    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not (AdoRs.BOF And AdoRs.EOF) Then
        arrayRecord = AdoRs.GetRows
    Else
        GoTo Error_Rtn
    End If

    AdoRs.Close

    Call subControlLock(arrayRecord, False, Mc1("iControl"), Mc1("rControl"))

    Set AdoRs = Nothing
    Set arrayRecord = Nothing

Error_Rtn:

    Set AdoRs = Nothing
    Set arrayRecord = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub subControlLock(ByVal vARRAY As Variant, ByVal bAllLock As Boolean, ByVal iCtrl As Collection, ByVal rCtrl As Collection)
    Dim icount       As Integer
    Dim iarrCOUNT    As Integer

    If bAllLock Then
        For icount = 1 To iCtrl.COUNT
            iCtrl.Item(icount).Visible = False
        Next
        For icount = 1 To rCtrl.COUNT
            rCtrl.Item(icount).Visible = False
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
        
        For icount = 1 To rCtrl.COUNT
            If rCtrl.Item(icount).Tag <> 99 And rCtrl.Item(icount).Tag <> "INS_EMP" Then

                    For iarrCOUNT = 0 To UBound(vARRAY, 1)

                        If Val(rCtrl.Item(icount).Tag) = Val(vARRAY(iarrCOUNT, 0)) Then
                            rCtrl.Item(icount).Visible = True
                            Exit For
                        Else
                            rCtrl.Item(icount).Visible = False
                        End If

                    Next

            End If
        Next
        
    End If


End Sub
'
Private Sub txt_SMP_CUT_LOC_LostFocus()
    Call Form_Ref
    
End Sub
'
Private Sub txt_SMP_NO_Change()
Dim sPROD_CD As String
    
    txt_smp_no_p.Text = txt_SMP_NO.Text
    
    sPROD_CD = SMP_PROD_Check(Trim(txt_smp_no_p.Text))
    
    If sPROD_CD = "ER" Then
        txt_SMP_CUT_LOC.Text = ""
    Else
        txt_SMP_CUT_LOC.Text = Find_SMP_LOC(Trim(txt_smp_no_p.Text))
        sAuthority = Ship_Input_AUTH(Trim(txt_smp_no_p.Text), sUserID, sOldAuthority)
        bExpo_SMP = Expo_SMP_Check(Trim(txt_smp_no_p.Text))
    End If

End Sub

