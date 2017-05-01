VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form Mainmenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MAIN SYSTEM"
   ClientHeight    =   6405
   ClientLeft      =   1785
   ClientTop       =   3765
   ClientWidth     =   12945
   Icon            =   "Mainmenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   12945
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic_otherday 
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   1680
      Picture         =   "Mainmenu.frx":0CCA
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   8
      Top             =   4380
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pic_today 
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2310
      Picture         =   "Mainmenu.frx":0F76
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1860
      Top             =   3840
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   13350
      Picture         =   "Mainmenu.frx":1222
      ScaleHeight     =   1770
      ScaleWidth      =   1905
      TabIndex        =   3
      Top             =   8730
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   13350
      Picture         =   "Mainmenu.frx":22B9
      ScaleHeight     =   1770
      ScaleWidth      =   1905
      TabIndex        =   2
      Top             =   8730
      Width           =   1905
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   1770
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4920
      Left            =   2430
      TabIndex        =   4
      Top             =   4770
      Width           =   11400
      _Version        =   393216
      _ExtentX        =   20108
      _ExtentY        =   8678
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "Mainmenu.frx":3285
      UserResize      =   0
      TextTip         =   3
   End
   Begin Threed.SSCommand cmd_aksystem 
      Height          =   600
      Left            =   900
      TabIndex        =   19
      Top             =   4980
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":388F
      Caption         =   "      生产管制"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_exit 
      Height          =   600
      Left            =   13890
      TabIndex        =   18
      Top             =   9840
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":3C9A
      Caption         =   "     Logout"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_aqsystem 
      Height          =   600
      Left            =   13680
      TabIndex        =   17
      Top             =   6030
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":40DF
      Caption         =   "      质量管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_ahsystem 
      Height          =   600
      Left            =   13710
      TabIndex        =   16
      Top             =   5430
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":458E
      Caption         =   "      发货管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_bgsystem 
      Height          =   600
      Left            =   13680
      TabIndex        =   15
      Top             =   4860
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":4A5F
      Caption         =   "      轧钢作业"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_afsystem 
      Height          =   600
      Left            =   13530
      TabIndex        =   14
      Top             =   3660
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":4E89
      Caption         =   "      炼钢作业"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_aesystem 
      Height          =   600
      Left            =   90
      TabIndex        =   13
      Top             =   4170
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":5B5B
      Caption         =   "      工序管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_acsystem 
      Height          =   600
      Left            =   90
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":6047
      Caption         =   "      进程管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_absystem 
      Height          =   600
      Left            =   60
      TabIndex        =   11
      Top             =   3090
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":6483
      Caption         =   "      订单管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_aasystem 
      Height          =   600
      Left            =   210
      TabIndex        =   10
      Top             =   5580
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":68CE
      Caption         =   "           销售生产计划"
      Alignment       =   4
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_azsystem 
      Height          =   600
      Left            =   13620
      TabIndex        =   9
      Top             =   4260
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":6D23
      Caption         =   "     系统管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1050
      Left            =   4905
      TabIndex        =   0
      Top             =   3510
      Visible         =   0   'False
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   1852
      _Version        =   196609
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel lblState 
         Height          =   285
         Left            =   90
         Top             =   675
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   503
         Caption         =   "文件下载中....!!"
         Alignment       =   1
         BackgroundStyle =   1
         BorderEffect    =   0
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel lblFileName 
         Height          =   285
         Left            =   90
         Top             =   90
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   503
         Caption         =   ""
         Alignment       =   1
         BackgroundStyle =   1
         BorderEffect    =   0
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin MSComctlLib.ProgressBar PrgDown 
         Height          =   315
         Left            =   135
         Negotiate       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   0
      Picture         =   "Mainmenu.frx":7182
      ScaleHeight     =   3030
      ScaleMode       =   0  'User
      ScaleWidth      =   12945
      TabIndex        =   5
      Top             =   0
      Width           =   12945
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "公知事项"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   810
         TabIndex        =   6
         Top             =   4140
         Width           =   7935
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   585
         Picture         =   "Mainmenu.frx":13B3E
         Top             =   4140
         Width           =   135
      End
   End
   Begin Threed.SSCommand cmd_cgsystem 
      Height          =   600
      Left            =   13530
      TabIndex        =   22
      Top             =   3090
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":13D26
      Caption         =   "     中板轧钢"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_total 
      Height          =   600
      Left            =   210
      TabIndex        =   23
      Top             =   6360
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8421504
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":14105
      Caption         =   "           综合生产管理"
      Alignment       =   4
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_husms 
      Height          =   600
      Left            =   210
      TabIndex        =   24
      Top             =   7050
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8421504
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":1455A
      Caption         =   "           中厚板炼钢"
      Alignment       =   4
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_humill 
      Height          =   600
      Left            =   210
      TabIndex        =   25
      Top             =   7740
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8421504
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":1522C
      Caption         =   "           中厚板轧钢"
      Alignment       =   4
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_banmill 
      Height          =   600
      Left            =   210
      TabIndex        =   26
      Top             =   8430
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8421504
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":15656
      Caption         =   "      中板轧钢"
      Alignment       =   1
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_fire 
      Height          =   600
      Left            =   210
      TabIndex        =   27
      Top             =   9120
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8421504
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":15A35
      Caption         =   "           热处理部分"
      Alignment       =   4
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_bksystem 
      Height          =   600
      Left            =   13620
      TabIndex        =   28
      Top             =   6630
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":15F62
      Caption         =   "      生产管制"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_cksystem 
      Height          =   600
      Left            =   13680
      TabIndex        =   29
      Top             =   7140
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":1636D
      Caption         =   "      生产管制"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_besystem 
      Height          =   600
      Left            =   13620
      TabIndex        =   30
      Top             =   7560
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":16778
      Caption         =   "      工序管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_cesystem 
      Height          =   600
      Left            =   13590
      TabIndex        =   31
      Top             =   7980
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":16C64
      Caption         =   "      工序管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_desystem 
      Height          =   600
      Left            =   12270
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":17150
      Caption         =   "      工序管理"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_dgsystem 
      Height          =   600
      Left            =   12330
      TabIndex        =   33
      Top             =   3960
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":1763C
      Caption         =   "      热处理作业"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_dksystem 
      Height          =   600
      Left            =   12270
      TabIndex        =   34
      Top             =   3720
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":17B69
      Caption         =   "      生产管制"
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmd_arsystem 
      Height          =   600
      Left            =   0
      TabIndex        =   35
      Top             =   4560
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      BackColor       =   16777215
      PictureMaskColor=   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mainmenu.frx":17F74
      Caption         =   "           ERP接口管理"
      Alignment       =   4
      PictureAlignment=   1
      ShapeSize       =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "南钢板材三级计算机系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   38.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   765
      Left            =   3435
      TabIndex        =   20
      Top             =   3330
      Width           =   9915
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "公知事项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   2835
      TabIndex        =   21
      Top             =   4455
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   2610
      Picture         =   "Mainmenu.frx":18085
      Top             =   4455
      Width           =   135
   End
End
Attribute VB_Name = "Mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Sc1 As New Collection             'Spread Collection
Dim Proc_Sc As New Collection         'Spread Struc Collection
Dim PrevRow As Long                   'Mouse previous Move Row
Dim lTmr_Cnt As Long                  'Timer Space count

Dim btotal As Boolean
Dim bhusms As Boolean
Dim bhumill As Boolean
Dim bbanmill As Boolean
Dim bfire As Boolean

Public sServerIP As String            'SERVER IP
Public sServerID As String            'SERVER ID
Public sServerPWD As String           'SERVER PASSWORD
Public sServerPATH As String          'SERVER PATH
Public FILE_SIZE As Double            'FILE SIZE
Public bUnload As Boolean

Private Sub Form_Define()
        
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Me.KeyPreview = True
    
    btotal = False
    bhusms = False
    bhumill = False
    bbanmill = False
    bfire = False
    
End Sub

Private Sub cmd_banmill_Click()

    Call Button_All_Visible
    
    cmd_total.ForeColor = &H808080
    cmd_husms.ForeColor = &H808080
    cmd_humill.ForeColor = &H808080
    cmd_banmill.ForeColor = &HFF&
    cmd_fire.ForeColor = &H808080
    
    btotal = False
    bhusms = False
    bhumill = False
    bbanmill = True
    bfire = False
    
    cmd_cesystem.Top = 9840
    cmd_cesystem.Left = 2085
    
    cmd_cksystem.Top = 9840
    cmd_cksystem.Left = cmd_cesystem.Left + cmd_cesystem.Width + 30
    
    cmd_cgsystem.Top = 9840
    cmd_cgsystem.Left = cmd_cksystem.Left + cmd_cksystem.Width + 30
    
    cmd_cesystem.Visible = True
    cmd_cksystem.Visible = True
    cmd_cgsystem.Visible = True
    
End Sub

Private Sub cmd_banmill_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If bbanmill Then
        cmd_banmill.ForeColor = &HFF0000
    Else
        cmd_banmill.ForeColor = &H808080
    End If
    
End Sub

Private Sub cmd_banmill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bbanmill Then
        cmd_banmill.ForeColor = &HFF0000
    Else
        cmd_banmill.ForeColor = &HFF&
    End If
    
End Sub

Private Sub cmd_aasystem_Click()
    Call ExeRun_File("AA.exe", "销售生产计划")
End Sub

Private Sub cmd_absystem_Click()
    Call ExeRun_File("AB.exe", "订单管理")
End Sub

Private Sub cmd_acsystem_Click()
    Call ExeRun_File("AC.exe", "进程管理")
End Sub

Private Sub cmd_aesystem_Click()
    Call ExeRun_File("AE.exe", "板卷炼钢工序管理")
End Sub

Private Sub cmd_besystem_Click()
    Call ExeRun_File("BE.exe", "板卷轧钢工序管理")
End Sub

Private Sub cmd_cesystem_Click()
    Call ExeRun_File("CE.exe", "中板轧钢工序管理")
End Sub

Private Sub cmd_desystem_Click()
    Call ExeRun_File("DE.exe", "热处理工序管理")
End Sub

Private Sub cmd_afsystem_Click()
    Call ExeRun_File("AF.exe", "板卷炼钢作业管理")
End Sub

Private Sub cmd_bgsystem_Click()
    Call ExeRun_File("BG.exe", "板卷轧钢作业管理")
End Sub

Private Sub cmd_cgsystem_Click()
    Call ExeRun_File("CG.exe", "中板轧钢作业管理")
End Sub

Private Sub cmd_dgsystem_Click()
    Call ExeRun_File("DG.exe", "热处理作业管理")
End Sub

Private Sub cmd_ahsystem_Click()
    Call ExeRun_File("AH.exe", "发货管理")
End Sub

Private Sub cmd_aksystem_Click()
    Call ExeRun_File("AK.exe", "炼钢生产管制")
End Sub

Private Sub cmd_bksystem_Click()
    Call ExeRun_File("BK.exe", "板卷生产管制")
End Sub

Private Sub cmd_cksystem_Click()
    Call ExeRun_File("CK.exe", "中板生产管制")
End Sub

Private Sub cmd_dksystem_Click()
    Exit Sub
    Call ExeRun_File("DK.exe", "热处理生产管制")
End Sub

Private Sub cmd_aqsystem_Click()
    Call ExeRun_File("AQ.exe", "质量管理")
End Sub

Private Sub cmd_azsystem_Click()
    Call ExeRun_File("AZ.exe", "系统管理")
End Sub

Private Sub cmd_arsystem_Click()
'    Exit Sub
    Call ExeRun_File("AR.exe", "ERP接口管理")
End Sub

Private Sub cmd_exit_Click()
    
    If bUnload Then
    
        'AA.exe Active
        If Process_Exe_Check("AA.exe") Then
            Call Gp_MsgBoxDisplay("销售生产计划...实行中...!!", "I")
            Exit Sub
        End If
        
        'AB.exe Active
        If Process_Exe_Check("AB.exe") Then
            Call Gp_MsgBoxDisplay("订单管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'AC.exe Active
        If Process_Exe_Check("AC.exe") Then
            Call Gp_MsgBoxDisplay("进程管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'AE.exe Active
        If Process_Exe_Check("AE.exe") Then
            Call Gp_MsgBoxDisplay("板卷炼钢工序管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'BE.exe Active
        If Process_Exe_Check("BE.exe") Then
            Call Gp_MsgBoxDisplay("板卷轧钢工序管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'CE.exe Active
        If Process_Exe_Check("CE.exe") Then
            Call Gp_MsgBoxDisplay("中板轧钢工序管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'DE.exe Active
        If Process_Exe_Check("DE.exe") Then
            Call Gp_MsgBoxDisplay("热处理工序管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'AF.exe Active
        If Process_Exe_Check("AF.exe") Then
            Call Gp_MsgBoxDisplay("板卷炼钢作业管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'BG.exe Active
        If Process_Exe_Check("BG.exe") Then
            Call Gp_MsgBoxDisplay("板卷轧钢作业管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'CG.exe Active
        If Process_Exe_Check("CG.exe") Then
            Call Gp_MsgBoxDisplay("中板轧钢作业管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'DG.exe Active
        If Process_Exe_Check("DG.exe") Then
            Call Gp_MsgBoxDisplay("热处理作业管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'AH.exe Active
        If Process_Exe_Check("AH.exe") Then
            Call Gp_MsgBoxDisplay("发货管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'AK.exe Active
        If Process_Exe_Check("AK.exe") Then
            Call Gp_MsgBoxDisplay("板卷炼钢生产管制...实行中...!!", "I")
            Exit Sub
        End If
        
        'BK.exe Active
        If Process_Exe_Check("BK.exe") Then
            Call Gp_MsgBoxDisplay("板卷轧钢生产管制...实行中...!!", "I")
            Exit Sub
        End If
        
        'CK.exe Active
        If Process_Exe_Check("CK.exe") Then
            Call Gp_MsgBoxDisplay("中板轧钢生产管制...实行中...!!", "I")
            Exit Sub
        End If
        
        'DK.exe Active
        If Process_Exe_Check("DK.exe") Then
            Call Gp_MsgBoxDisplay("热处理生产管制...实行中...!!", "I")
            Exit Sub
        End If
        
        'AQ.exe Active
        If Process_Exe_Check("AQ.exe") Then
            Call Gp_MsgBoxDisplay("质量管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'AZ.exe Active
        If Process_Exe_Check("AZ.exe") Then
            Call Gp_MsgBoxDisplay("系统管理...实行中...!!", "I")
            Exit Sub
        End If
        
        'AR.exe Active
        If Process_Exe_Check("AR.exe") Then
            Call Gp_MsgBoxDisplay("ERP接口管理...实行中...!!", "I")
            Exit Sub
        End If
        
    End If
    
    cmd_aasystem.Visible = False
    cmd_absystem.Visible = False
    cmd_acsystem.Visible = False
    cmd_aesystem.Visible = False
    cmd_besystem.Visible = False
    cmd_cesystem.Visible = False
    cmd_desystem.Visible = False
    cmd_aksystem.Visible = False
    cmd_bksystem.Visible = False
    cmd_cksystem.Visible = False
    cmd_dksystem.Visible = False
    cmd_afsystem.Visible = False
    cmd_bgsystem.Visible = False
    cmd_cgsystem.Visible = False
    cmd_dgsystem.Visible = False
    cmd_ahsystem.Visible = False
    cmd_aqsystem.Visible = False
    cmd_arsystem.Visible = False
    cmd_azsystem.Visible = False
    
    cmd_total.Visible = False
    cmd_husms.Visible = False
    cmd_humill.Visible = False
    cmd_banmill.Visible = False
    cmd_fire.Visible = False
    cmd_exit.Visible = False
    
    PassCheck = False
    Picture1.Visible = True
    Picture2.Visible = False
    Label1.Visible = True
    Image1.Visible = True
    ss1.Visible = True
    cmd_exit.Visible = False
    
    SaveSetting "NISCO", "AUTHORITY", "sUserID", ""
    SaveSetting "NISCO", "AUTHORITY", "sUsername", ""
    
    Call Sp_Display
    
End Sub

Private Sub cmd_fire_Click()

    Call Button_All_Visible
    
    cmd_total.ForeColor = &H808080
    cmd_husms.ForeColor = &H808080
    cmd_humill.ForeColor = &H808080
    cmd_banmill.ForeColor = &H808080
    cmd_fire.ForeColor = &HFF&
    
    btotal = False
    bhusms = False
    bhumill = False
    bbanmill = False
    bfire = True
    
    cmd_desystem.Top = 9840
    cmd_desystem.Left = 2085
    
'    cmd_dksystem.Top = 9840
'    cmd_dksystem.Left = cmd_desystem.Left + cmd_desystem.Width + 30
    
    cmd_dgsystem.Top = 9840
    cmd_dgsystem.Left = cmd_desystem.Left + cmd_desystem.Width + 30
    
    cmd_desystem.Visible = True
    'cmd_dksystem.Visible = True
    cmd_dgsystem.Visible = True
    
End Sub

Private Sub cmd_fire_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If bfire Then
        cmd_fire.ForeColor = &HFF0000
    Else
        cmd_fire.ForeColor = &H808080
    End If
    
End Sub

Private Sub cmd_fire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bfire Then
        cmd_fire.ForeColor = &HFF0000
    Else
        cmd_fire.ForeColor = &HFF&
    End If
    
End Sub

Private Sub cmd_humill_Click()

    Call Button_All_Visible
    
    cmd_total.ForeColor = &H808080
    cmd_husms.ForeColor = &H808080
    cmd_humill.ForeColor = &HFF&
    cmd_banmill.ForeColor = &H808080
    cmd_fire.ForeColor = &H808080
    
    btotal = False
    bhusms = False
    bhumill = True
    bbanmill = False
    bfire = False
    
    cmd_besystem.Top = 9840
    cmd_besystem.Left = 2085
    
    cmd_bksystem.Top = 9840
    cmd_bksystem.Left = cmd_besystem.Left + cmd_besystem.Width + 30
    
    cmd_bgsystem.Top = 9840
    cmd_bgsystem.Left = cmd_bksystem.Left + cmd_bksystem.Width + 30
    
    cmd_besystem.Visible = True
    cmd_bksystem.Visible = True
    cmd_bgsystem.Visible = True
    
End Sub

Private Sub cmd_humill_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If bhumill Then
        cmd_humill.ForeColor = &HFF0000
    Else
        cmd_humill.ForeColor = &H808080
    End If
'
End Sub

Private Sub cmd_humill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bhumill Then
        cmd_humill.ForeColor = &HFF0000
    Else
        cmd_humill.ForeColor = &HFF&
    End If
    
End Sub

Private Sub cmd_husms_Click()

    Call Button_All_Visible
    
    cmd_total.ForeColor = &H808080
    cmd_husms.ForeColor = &HFF&
    cmd_humill.ForeColor = &H808080
    cmd_banmill.ForeColor = &H808080
    cmd_fire.ForeColor = &H808080
    
    btotal = False
    bhusms = True
    bhumill = False
    bbanmill = False
    bfire = False
    
    cmd_aesystem.Top = 9840
    cmd_aesystem.Left = 2085
    
    cmd_aksystem.Top = 9840
    cmd_aksystem.Left = cmd_aesystem.Left + cmd_aesystem.Width + 30
    
    cmd_afsystem.Top = 9840
    cmd_afsystem.Left = cmd_aksystem.Left + cmd_aksystem.Width + 30
    
    cmd_aesystem.Visible = True
    cmd_aksystem.Visible = True
    cmd_afsystem.Visible = True
    
End Sub

Private Sub cmd_husms_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If bhusms Then
        cmd_husms.ForeColor = &HFF0000
    Else
        cmd_husms.ForeColor = &H808080
    End If
    
End Sub

Private Sub cmd_husms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bhusms Then
        cmd_husms.ForeColor = &HFF0000
    Else
        cmd_husms.ForeColor = &HFF&
    End If
    
End Sub

Private Sub cmd_total_Click()

    Call Button_All_Visible

    cmd_total.ForeColor = &HFF&
    cmd_husms.ForeColor = &H808080
    cmd_humill.ForeColor = &H808080
    cmd_banmill.ForeColor = &H808080
    cmd_fire.ForeColor = &H808080
    
    btotal = True
    bhusms = False
    bhumill = False
    bbanmill = False
    bfire = False
    
    cmd_aasystem.Top = 9840
    cmd_aasystem.Left = 2085
    
    cmd_absystem.Top = 9840
    cmd_absystem.Left = cmd_aasystem.Left + cmd_aasystem.Width + 30
    
    cmd_aqsystem.Top = 9840
    cmd_aqsystem.Left = cmd_absystem.Left + cmd_absystem.Width + 30
    
    cmd_acsystem.Top = 9840
    cmd_acsystem.Left = cmd_aqsystem.Left + cmd_aqsystem.Width + 30
    
    cmd_ahsystem.Top = 9840
    cmd_ahsystem.Left = cmd_acsystem.Left + cmd_acsystem.Width + 30
    
    cmd_arsystem.Top = 9840
    cmd_arsystem.Left = cmd_ahsystem.Left + cmd_ahsystem.Width + 30
    
    cmd_azsystem.Top = 9840
    cmd_azsystem.Left = cmd_arsystem.Left + cmd_arsystem.Width + 30
    
    cmd_aasystem.Visible = True
    cmd_absystem.Visible = True
    cmd_aqsystem.Visible = True
    cmd_acsystem.Visible = True
    cmd_ahsystem.Visible = True
    cmd_arsystem.Visible = True
    cmd_azsystem.Visible = True
    
End Sub

Private Sub cmd_total_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If btotal Then
        cmd_total.ForeColor = &HFF0000
    Else
        cmd_total.ForeColor = &H808080
    End If
    
End Sub

Private Sub cmd_total_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If btotal Then
        cmd_total.ForeColor = &HFF0000
    Else
        cmd_total.ForeColor = &HFF&
    End If
    
End Sub

Private Sub Form_Load()

On Error GoTo Find_Error

    Dim X As Boolean
    
    Dim Active_YN As String
    Dim AdoRs As ADODB.Recordset
    Set AdoRs = New ADODB.Recordset
    
    bUnload = True
    Active_YN = GetSetting("NISCO", "EXE-FILE", "Main.exe")

    If Active_YN = "1" Then
        SaveSetting "NISCO", "EXE-FILE", "Main.exe", ""
    Else
        bUnload = False
        Call Gp_MsgBoxDisplay("Nisco...Exectue", "W")
        Unload Me
        Exit Sub
    End If

    Me.Show
    PassCheck = False
    
    If GF_DbConnect = False Then Unload Me
    Picture1.Visible = True
    Picture2.Visible = False
    
    Call Form_Define
    Call Sp_Setting
    Call Sp_ReadOnlySet
    Call Sp_Cls
    Call Sp_Display
    
    X = ss1.SetTextTipAppearance("SimSun", "10", False, False, &HC0FFFF, &H800000)
    
    M_CN1.Close
    Set M_CN1 = Nothing

    Exit Sub
    
Find_Error:
    Set AdoRs = Nothing
    
End Sub

Public Function GetFileName(Conn As ADODB.Connection, lsFileLike As String, ByRef sFileName() As String) As Variant

On Error GoTo FloatFind_Error

    Dim AdoRs As ADODB.Recordset
    Dim sQuery As String
    Dim lnCount As Integer
    
    Set AdoRs = New ADODB.Recordset
    
    GetFileName = 0
    sQuery = "SELECT REPORTNAME FROM ZP_REPORTNAME WHERE SUB_SYS LIKE '%" + Mid(lsFileLike, 2, 5) + "%'"
    
    'Update 2007.06.08
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then GetFileName = 0: Exit Function
    End If
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    lnCount = 0
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        ReDim sFileName(AdoRs.RecordCount)
        GetFileName = AdoRs.RecordCount
        While Not AdoRs.EOF
            lnCount = lnCount + 1
            If VarType(AdoRs.Fields(0)) = vbNull Then
                sFileName(lnCount) = ""
            Else
                sFileName(lnCount) = AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
        Wend
        
    Else
        GetFileName = 0
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

FloatFind_Error:

    Set AdoRs = Nothing
    GetFileName = 0

End Function

Private Sub ExeRun_File(WinId As String, Form_Caption As String)
    
On Error GoTo Err_Handler

    Dim sQuery As String
    Dim Active_YN As String
    Dim Client_Ver As String
    Dim Server_Ver As String
    Dim sFilePath As String
    
    Dim sTnsPath As String
    Dim sTnsPath_Start As Integer
    
    Dim lHandle As Long
    Dim lFileSize As Long
    Dim I, lnCount As Integer
    Dim sDownLoadFile() As String
    
    PrgDown.Value = 0
    
    Call Server_Info
    
    If sServerIP = "" Or sServerID = "" Or sServerPWD = "" Or sServerPATH = "" Then
        Call Gp_MsgBoxDisplay("服务器相关信息不正确...!!")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Active_YN = GetSetting("NISCO", "EXE-FILE", WinId)
    
    If Process_Exe_Check(WinId) Then
        Call Gp_MsgBoxDisplay("已在运行" + Form_Caption, "I")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'If Active_YN = "1" Then
        'Screen.MousePointer = vbDefault
        'Exit Sub
    'End If
    
    'FILE SIZE
    FILE_SIZE = Gf_FloatFind(M_CN1, "SELECT SYS_SIZE FROM ZP_VERSION WHERE SUB_SYS = '" & WinId & "' ")
    
    'Server Version
    sQuery = "SELECT TRIM(FST_VER) || TRIM(SND_VER) || TRIM(THR_VER) FROM ZP_VERSION WHERE SUB_SYS = '" & WinId & "' "
    Server_Ver = Gf_CodeFind(M_CN1, sQuery)
    
    'Client Version
    Client_Ver = Trim(Str(GetPrivateProfileInt("Version", WinId, 0, App.Path & "\" & "ENVI.INI")))
    
    Client_Ver = "000000" & Client_Ver
    
    Client_Ver = Right(Client_Ver, 6)
    
    With Inet
    
        Inet.Cancel
        
        If InStr(1, Server_Ver, Client_Ver, vbTextCompare) = 0 Or Dir(App.Path & "\" & WinId) = "" Then

            lnCount = GetFileName(M_CN1, WinId, sDownLoadFile)
            lblFileName.Caption = ""
            SSPanel1.Visible = True
            
            'DownLoad ReportFile
            For I = 1 To lnCount
            
                If sDownLoadFile(I) <> "" Then
                    
                    If Dir(App.Path & "\" & sDownLoadFile(I)) <> "" Then
                        Kill App.Path & "\" & sDownLoadFile(I)
                    End If
                    
                    .Execute , "GET " & sServerPATH & sDownLoadFile(I) & " " & Chr(34) & App.Path & "\" & sDownLoadFile(I) & Chr(34)
                
                    Do While .StillExecuting
                        DoEvents
                    Loop
                    
                End If
                
            Next I
            
            '.Execute , "quit"
            
            PrgDown.Max = FILE_SIZE
            PrgDown.Value = 0
            
            'Client File Delete
            If Dir(App.Path & "\" & WinId) <> "" Then
                Kill App.Path & "\" & WinId
            End If
            
            'Server -> Client Copy
            .Execute , "GET " & sServerPATH & WinId & " " & Chr(34) & App.Path & "\" & WinId & Chr(34)
            
            Do While .StillExecuting
            
                Sleep (200)  'add yangmeng at 091116
                DoEvents
                
'                lblFileName.Caption = Format(FileLen(App.Path & "\" & WinId) \ 1024, "#,##0") & " KB" & " / " & Format(FILE_SIZE \ 1024, "#,##0") & " KB"
'
'                If FileLen(App.Path & "\" & WinId) > FILE_SIZE Then
'                    PrgDown.Value = FILE_SIZE
'                Else
'                    PrgDown.Value = FileLen(App.Path & "\" & WinId)
'                End If
            Loop
            
            lblFileName.Caption = Format(FileLen(App.Path & "\" & WinId) \ 1024, "#,##0") & " KB" & " / " & Format(FILE_SIZE \ 1024, "#,##0") & " KB"
            
            If FileLen(App.Path & "\" & WinId) > FILE_SIZE Then
                PrgDown.Value = FILE_SIZE
            Else
                PrgDown.Value = FileLen(App.Path & "\" & WinId)
            End If
            
            'END
            .Execute , "quit"
            
            Do While .StillExecuting
                DoEvents
            Loop
            
            Call WritePrivateProfileString("Version", WinId, Server_Ver, App.Path & "\" & "ENVI.INI")
            
            .Cancel
            Do While .StillExecuting
                DoEvents
            Loop
            
            SSPanel1.Visible = False
            
        End If
    
    End With
    
    'CALL
    SaveSetting "NISCO", "EXE-FILE", WinId, "1"
    Shell App.Path & "\" & WinId, vbMaximizedFocus
    
    I = 0
    While FindWindow(vbNullString, Form_Caption) = 0 And I < 3
      Sleep (700)
      I = I + 1
    Wend

    I = 0
    While GetSetting("NISCO", "EXE-FILE", WinId) <> "ok" And I < 25
        Sleep (100)
        I = I + 1
    Wend
    
    SaveSetting "NISCO", "EXE-FILE", WinId, ""
    Screen.MousePointer = vbDefault
    
    'Update 2007.06.08
    M_CN1.Close
    Set M_CN1 = Nothing
    
    Exit Sub
    
Err_Handler:

    M_CN1.Close
    Set M_CN1 = Nothing
    SSPanel1.Visible = False
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If PassCheck Then
        Picture1.Visible = False
        Picture2.Visible = False
    Else
        Picture1.Visible = True
        Picture2.Visible = False
        Picture1.Drag vbEndDrag
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If bUnload Then
    
        'AA.exe Active
        If Process_Exe_Check("AA.exe") Then
            Call Gp_MsgBoxDisplay("销售生产计划...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AB.exe Active
        If Process_Exe_Check("AB.exe") Then
            Call Gp_MsgBoxDisplay("订单管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AC.exe Active
        If Process_Exe_Check("AC.exe") Then
            Call Gp_MsgBoxDisplay("进程管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AE.exe Active
        If Process_Exe_Check("AE.exe") Then
            Call Gp_MsgBoxDisplay("板卷炼钢工序管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'BE.exe Active
        If Process_Exe_Check("BE.exe") Then
            Call Gp_MsgBoxDisplay("板卷轧钢工序管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'CE.exe Active
        If Process_Exe_Check("CE.exe") Then
            Call Gp_MsgBoxDisplay("中板轧钢工序管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'DE.exe Active
        If Process_Exe_Check("DE.exe") Then
            Call Gp_MsgBoxDisplay("热处理工序管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AF.exe Active
        If Process_Exe_Check("AF.exe") Then
            Call Gp_MsgBoxDisplay("板卷炼钢作业管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'BG.exe Active
        If Process_Exe_Check("BG.exe") Then
            Call Gp_MsgBoxDisplay("板卷轧钢作业管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'CG.exe Active
        If Process_Exe_Check("CG.exe") Then
            Call Gp_MsgBoxDisplay("中板轧钢作业管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'DG.exe Active
        If Process_Exe_Check("DG.exe") Then
            Call Gp_MsgBoxDisplay("热处理作业管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AH.exe Active
        If Process_Exe_Check("AH.exe") Then
            Call Gp_MsgBoxDisplay("发货管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AK.exe Active
        If Process_Exe_Check("AK.exe") Then
            Call Gp_MsgBoxDisplay("板卷炼钢生产管制...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'BK.exe Active
        If Process_Exe_Check("BK.exe") Then
            Call Gp_MsgBoxDisplay("板卷轧钢生产管制...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'CK.exe Active
        If Process_Exe_Check("CK.exe") Then
            Call Gp_MsgBoxDisplay("中板轧钢生产管制...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'DK.exe Active
        If Process_Exe_Check("DK.exe") Then
            Call Gp_MsgBoxDisplay("热处理生产管制...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AQ.exe Active
        If Process_Exe_Check("AQ.exe") Then
            Call Gp_MsgBoxDisplay("质量管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AZ.exe Active
        If Process_Exe_Check("AZ.exe") Then
            Call Gp_MsgBoxDisplay("系统管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        'AR.exe Active
        If Process_Exe_Check("AR.exe") Then
            Call Gp_MsgBoxDisplay("ERP接口管理...实行中...!!", "I")
            Cancel = True
            Exit Sub
        End If
        
        SaveSetting "NISCO", "AUTHORITY", "sUserID", ""
        SaveSetting "NISCO", "AUTHORITY", "sUsername", ""
        
        If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "Main.exe", ""
        
    End If
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Picture1.Visible = False
    Picture2.Visible = True
    Picture2.Drag vbEndDrag

End Sub

Private Sub Picture2_Click()

    Login.Show 1
    
    If PassCheck Then
    
        Picture1.Visible = False
        cmd_total.Visible = True
        cmd_husms.Visible = True
        cmd_humill.Visible = True
        cmd_banmill.Visible = True
        cmd_fire.Visible = True
        cmd_exit.Visible = True
        
        cmd_total.ForeColor = &H808080
        cmd_husms.ForeColor = &H808080
        cmd_humill.ForeColor = &H808080
        cmd_banmill.ForeColor = &H808080
        cmd_fire.ForeColor = &H808080
        
        Picture1.Visible = False
        Picture2.Visible = False

        Label1.Visible = False
        Image1.Visible = False
        'ss1.Visible = False
        
    End If

End Sub

Private Sub Sp_Setting()

    With ss1
    
        .RowHeight(-1) = 13
        .BackColorStyle = BackColorStyleUnderGrid
        
        .BorderStyle = BorderStyleNone
        .RowHeadersShow = False
        .ColHeadersShow = False
        
        .GrayAreaBackColor = &HFFFFFF
        .GridColor = &HFFFFFF
        .ShadowColor = &HFFFFFF
        .ShadowDark = &HFFFFFF
        .SelBackColor = &HFFFFFF        ''&HE3F4FF      ''&HFFFF80     '&H808040

        .OperationMode = OperationModeRead
        .UserResize = UserResizeNone

        .ProcessTab = True
        .ScrollBarExtMode = True
        .ScrollBars = ScrollBarsVertical
        .ButtonDrawMode = 1
        .TabStop = False

        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1

        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False

        .MaxRows = 0
        
    End With
    
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If PassCheck Then
        Picture1.Visible = False
        Picture2.Visible = False
    Else
        Picture1.Visible = True
        Picture2.Visible = False
        Picture1.Drag vbEndDrag
    End If
    
End Sub

Private Sub ss1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim Col As Long, Row As Long

    ss1.GetCellFromScreenCoord Col, Row, X, Y
    If Row <= 0 Or PrevRow = Row Then Exit Sub

    Call Sp_RowColor(ss1, Row, , &HFFE0D7)
    Call Sp_RowColor(ss1, PrevRow)
    PrevRow = Row

    If PassCheck Then
        Picture1.Visible = False
        Picture2.Visible = False
    Else
        Picture1.Visible = True
        Picture2.Visible = False
        Picture1.Drag vbEndDrag
    End If
    
End Sub

Private Sub ss1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim sString As String
    
    If Col <> 2 Then Exit Sub
    
    ss1.Col = Col
    ss1.Row = Row
    
    If Len(ss1.Text) > 60 Then
        ShowTip = True
        TipText = ss1.Text
    End If
    
End Sub

Private Sub Sp_ReadOnlySet()

    With ss1
    
        .Col = 0: .Col2 = .MaxCols
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .Lock = True
        .BlockMode = False
    
    End With
    
End Sub

Private Sub Sp_Cls()

    With ss1
        
        .MaxRows = 0
        .OperationMode = OperationModeNormal
        
    End With

End Sub

Private Sub Sp_RowColor(sPname As Variant, iRow As Variant, Optional fColor As Variant = vbBlack, _
                          Optional bColor As Variant = vbWhite)

    With sPname

        .Col = 1: .Col2 = -1
        .Row = iRow: .Row2 = iRow
        
        .BlockMode = True
        .ForeColor = fColor
        .BackColor = bColor
        .BlockMode = False

    End With

End Sub

Private Sub Server_Info()

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    Set AdoRs = New ADODB.Recordset
    
    sQuery = "SELECT SERVER_IP, SERVER_ID, SERVER_PWD, SERVER_PATH FROM ZP_SERVERINFO "
    
    'Update 2007.06.08
    'Db Connection Check
    If M_CN1.State = 0 Then
        If GF_DbConnect = False Then Exit Sub
    End If
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not AdoRs.BOF And Not AdoRs.EOF Then

        If VarType(AdoRs.Fields(0)) = vbNull Then
            sServerIP = ""
        Else
            sServerIP = AdoRs.Fields(0)
        End If

        If VarType(AdoRs.Fields(1)) = vbNull Then
            sServerID = ""
        Else
            sServerID = AdoRs.Fields(1)
        End If

        If VarType(AdoRs.Fields(2)) = vbNull Then
            sServerPWD = ""
        Else
            sServerPWD = AdoRs.Fields(2)
        End If

        If VarType(AdoRs.Fields(3)) = vbNull Then
            sServerPATH = ""
        Else
            sServerPATH = AdoRs.Fields(3)
        End If

    End If

    AdoRs.Close
    Set AdoRs = Nothing
    
    Inet.Protocol = icFTP
    Inet.URL = sServerIP
    Inet.UserName = sServerID
    Inet.Password = sServerPWD
    
End Sub

Private Sub Sp_Display()

    Dim sQuery As String
    
    sQuery = "SELECT '', SCRIPT, SUBSTR(INS_DATE,3,6)||INS_TIME, GF_EMPNAMEFIND(INS_EMP) FROM ZP_INFORMATION "
    sQuery = sQuery + " WHERE ROWNUM < 16 ORDER BY INS_SEQ DESC "
    
    If Not Only_Display(M_CN1, Proc_Sc("Sc"), sQuery, , False, False) Then
        Image1.Visible = False
        Label1.Visible = False
    End If
    
    'Update 2007.06.08
    M_CN1.Close
    Set M_CN1 = Nothing

End Sub

Private Function Only_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional iDupCnt As Variant = 0, _
                                Optional MsgChk As Boolean = True, Optional EvenRowChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn
    
    Dim JJ, j As Integer
    
    Dim lRowCount As Long
    Dim lColCount As Long
    Dim sTemp() As String
    Dim sToday As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    'Update 2007.0.08
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Only_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    sToday = Gf_CodeFind(Conn, "SELECT TO_CHAR(SYSDATE,'YYMMDD') FROM DUAL")
        
    With Sc.Item("Spread")

        Only_Display = True
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
            Only_Display = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = 0
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing
        
        If iDupCnt > 0 Then
            ReDim sTemp(0 To iDupCnt - 1)
        End If
        
        If UBound(ArrayRecords, 1) >= 0 Then
        
            .MaxRows = UBound(ArrayRecords, 2) + 1
        
            For lRowCount = 0 To .MaxRows - 1
            
                .Row = lRowCount + 1
                
                'Duplicate Process
                For j = 1 To iDupCnt Step 1
                
                    If sTemp(j - 1) <> Trim(ArrayRecords(j - 1, lRowCount)) Then
                        .Col = j
                        .Text = Trim(ArrayRecords(j - 1, lRowCount))
                        sTemp(j - 1) = Trim(ArrayRecords(j - 1, lRowCount))
                        
                        For JJ = j + 1 To iDupCnt Step 1
                            sTemp(JJ - 1) = ""
                        Next JJ
                        
                    End If
                    
                Next j
            
                For lColCount = iDupCnt To .MaxCols - 1
                
                    .Col = lColCount + 1
                    
                    If .Col = 1 Then
                        .TypePictCenter = True
                        .TypePictMaintainScale = True
                        .TypePictStretch = False
                        
                        Select Case Mid(Trim(ArrayRecords(2, lRowCount)), 1, 6)
                            Case sToday
                                .TypePictPicture = pic_today.Picture
                            Case Else
                                .TypePictPicture = pic_otherday.Picture
                        End Select
                    
                    ElseIf .Col = 3 Then   'SS_CELL_TYPE_PIC (DATE) --> VALUE
                        If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                            .Value = ""
                        Else
                            .Value = Trim(ArrayRecords(lColCount, lRowCount))
                        End If
                    Else
                        If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(lColCount, lRowCount))
                        End If
                    End If
                    
                Next lColCount
                
            Next lRowCount
        End If
                                            
        .ReDraw = True
        
    End With
    
    Sc.Item("Spread").OperationMode = OperationModeRead
    Only_Display = True
    Screen.MousePointer = vbDefault
    Exit Function
   
Error_Rtn:

    Set AdoRs = Nothing
    Only_Display = False

    Screen.MousePointer = vbDefault
    
End Function

Private Sub Button_All_Visible()

    cmd_aasystem.Visible = False
    cmd_absystem.Visible = False
    cmd_acsystem.Visible = False
    cmd_aesystem.Visible = False
    cmd_besystem.Visible = False
    cmd_cesystem.Visible = False
    cmd_desystem.Visible = False
    cmd_aksystem.Visible = False
    cmd_bksystem.Visible = False
    cmd_cksystem.Visible = False
    cmd_dksystem.Visible = False
    cmd_afsystem.Visible = False
    cmd_bgsystem.Visible = False
    cmd_cgsystem.Visible = False
    cmd_dgsystem.Visible = False
    cmd_ahsystem.Visible = False
    cmd_aqsystem.Visible = False
    cmd_arsystem.Visible = False
    cmd_azsystem.Visible = False

End Sub

Private Function Process_Exe_Check(Exe_File As String) As Boolean

    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim FileName  As String
    Dim r         As Long
    Dim lnghProcess As Long

    Process_Exe_Check = False
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)

    Do While r

        FileName = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))

        If UCase(Exe_File) = UCase(FileName) Then
            Process_Exe_Check = True
            Exit Function
        End If
        
        r = Process32Next(hSnapShot, uProcess)

    Loop

End Function
