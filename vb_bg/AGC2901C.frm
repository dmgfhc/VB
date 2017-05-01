VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2901C 
   Caption         =   "综合查询_AGC2901C"
   ClientHeight    =   9435
   ClientLeft      =   1005
   ClientTop       =   1605
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   15240
   Tag             =   "O.STLGRD"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   11280
      TabIndex        =   22
      Top             =   90
      Width           =   3885
      Begin VB.TextBox TXT_PROD_CD 
         Height          =   270
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin Threed.SSOption opt_Product 
         Height          =   330
         Index           =   0
         Left            =   480
         TabIndex        =   26
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "全部"
      End
      Begin Threed.SSOption opt_Product 
         Height          =   330
         Index           =   1
         Left            =   1620
         TabIndex        =   27
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "钢板"
      End
      Begin Threed.SSOption opt_Product 
         Height          =   330
         Index           =   2
         Left            =   2760
         TabIndex        =   28
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "钢卷"
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   11130
      Begin VB.TextBox TXT_SP_CD 
         Height          =   270
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox txt_Group_CD 
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
         ItemData        =   "AGC2901C.frx":0000
         Left            =   7575
         List            =   "AGC2901C.frx":0010
         TabIndex        =   2
         Top             =   260
         Width           =   735
      End
      Begin VB.ComboBox txt_Shift 
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
         ItemData        =   "AGC2901C.frx":0020
         Left            =   5730
         List            =   "AGC2901C.frx":002D
         TabIndex        =   1
         Top             =   260
         Width           =   735
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   480
         Top             =   260
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "生产日期"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   4860
         Top             =   255
         Width           =   855
         _ExtentX        =   1508
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   6705
         Top             =   255
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "班别"
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
      Begin InDate.UDate txt_DateFrom 
         Height          =   315
         Left            =   1785
         TabIndex        =   3
         Top             =   260
         Width           =   1395
         _ExtentX        =   2461
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
      Begin InDate.UDate txt_DateTo 
         Height          =   315
         Left            =   3195
         TabIndex        =   4
         Top             =   260
         Width           =   1395
         _ExtentX        =   2461
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
      Begin Threed.SSOption OPT_SLAB 
         Height          =   330
         Left            =   9900
         TabIndex        =   24
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "轧制"
      End
      Begin Threed.SSOption OPT_PLATE 
         Height          =   330
         Left            =   8820
         TabIndex        =   25
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "剪切"
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1155
      Left            =   180
      TabIndex        =   5
      Top             =   780
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   2037
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "轧制钢种"
         Height          =   255
         Index           =   16
         Left            =   1520
         TabIndex        =   40
         Tag             =   ",O.STDSPEC"
         Top             =   810
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "成品钢种"
         Height          =   255
         Index           =   15
         Left            =   270
         TabIndex        =   39
         Tag             =   ",O.STLGRD"
         Top             =   810
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "订单序列号"
         Height          =   255
         Index           =   14
         Left            =   6030
         TabIndex        =   38
         Tag             =   ",B.ORD_ITEM"
         Top             =   780
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "订单材代码"
         Height          =   255
         Index           =   13
         Left            =   7770
         TabIndex        =   37
         Tag             =   ",B.ORD_FL"
         Top             =   165
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "订单号"
         Height          =   255
         Index           =   12
         Left            =   6520
         TabIndex        =   36
         Tag             =   ",B.ORD_NO"
         Top             =   165
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "切边"
         Height          =   255
         Index           =   11
         Left            =   7770
         TabIndex        =   35
         Tag             =   ",B.TRIM_FL"
         Top             =   480
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "定尺"
         Height          =   255
         Index           =   10
         Left            =   6520
         TabIndex        =   34
         Tag             =   ",B.SIZE_KND"
         Top             =   480
         Width           =   1230
      End
      Begin VB.TextBox txt_Defect_name 
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
         Left            =   11040
         TabIndex        =   31
         Top             =   720
         Width           =   2475
      End
      Begin VB.TextBox txt_Defect 
         Alignment       =   2  'Center
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
         Left            =   13530
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txt_Disp_Order 
         Height          =   510
         Left            =   9420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   150
         Width           =   4785
      End
      Begin VB.TextBox txt_Order 
         Enabled         =   0   'False
         Height          =   510
         Left            =   11070
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   -480
         Width           =   3795
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "班别"
         Height          =   255
         Index           =   9
         Left            =   5270
         TabIndex        =   15
         Tag             =   ",B.GROUP_CD"
         Top             =   480
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "生产日"
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   17
         Tag             =   ",B.PROD_DATE"
         Top             =   480
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板坯号"
         Height          =   255
         Index           =   8
         Left            =   4020
         TabIndex        =   16
         Tag             =   ",substr(B.plate_no,1,10)"
         Top             =   480
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "成品标准"
         Height          =   255
         Index           =   1
         Left            =   1520
         TabIndex        =   14
         Tag             =   ",B.APLY_STDSPEC"
         Top             =   165
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "成品厚度"
         Height          =   255
         Index           =   2
         Left            =   2770
         TabIndex        =   13
         Tag             =   ",B.THK"
         Top             =   165
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "成品宽度"
         Height          =   255
         Index           =   3
         Left            =   4020
         TabIndex        =   12
         Tag             =   ",B.WID"
         Top             =   165
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "成品长度"
         Height          =   255
         Index           =   4
         Left            =   5270
         TabIndex        =   11
         Tag             =   ",B.LEN"
         Top             =   165
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板坯钢种"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Tag             =   ",B.STLGRD"
         Top             =   150
         Width           =   1230
      End
      Begin VB.TextBox txt_Disp 
         Height          =   345
         Left            =   14070
         TabIndex        =   9
         Top             =   210
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "入库日"
         Height          =   255
         Index           =   6
         Left            =   1520
         TabIndex        =   7
         Tag             =   ",SUBSTR(B.BED_PILE_DATE,1,8)"
         Top             =   480
         Width           =   1230
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "综合判定日"
         Height          =   255
         Index           =   7
         Left            =   2770
         TabIndex        =   6
         Tag             =   ",B.DSC_DATE"
         Top             =   480
         Width           =   1230
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   9420
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "缺陷"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7290
      Left            =   150
      TabIndex        =   18
      Top             =   2040
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12859
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "汇总信息"
      TabPicture(0)   =   "AGC2901C.frx":003A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ss1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "详细信息"
      TabPicture(1)   =   "AGC2901C.frx":0056
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ss2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "成材率"
      TabPicture(2)   =   "AGC2901C.frx":0072
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ss3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "设计成材率"
      TabPicture(3)   =   "AGC2901C.frx":008E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ss4"
      Tab(3).ControlCount=   1
      Begin FPSpread.vaSpread ss2 
         Height          =   6750
         Left            =   -74880
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   420
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   11906
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         ColsFrozen      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   33
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2901C.frx":00AA
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   6750
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   450
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   11906
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   15
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2901C.frx":42BC
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   6750
         Left            =   -74880
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   420
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   11906
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   7
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2901C.frx":616D
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   6750
         Left            =   -74880
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   420
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   11906
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   3
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2901C.frx":7C66
      End
   End
End
Attribute VB_Name = "AGC2901C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PLATE/COIL STOCK MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        AGC2901C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder             Yang meng
'-- Date              2005.11.11
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

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread Necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection


Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iSumCol   As New Collection       'Sum Column
Const iss1MaxCols = 15
Const iss1Ok_rat = 12
Const iss1Prod_rat = 13
Const iss1Plan_rat = 14
Const iss1Cut_wgt = 15
Const iss4MaxCols = 3

Const SS2_SLAB_NO = 1
Const SS2_MAT_NO = 2
Const SS2_ORD_NO = 12
Const SS2_ORD_ITEM = 13
Const SS2_URGNT_FL = 33



Private Sub Form_Define()

    Dim iIndex As Integer
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DateFrom, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_DateTo, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Shift, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_Group_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_SP_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Order, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_Defect, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
      
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '是否紧急订单
      
   Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   
   Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)

    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2901C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGC2901C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AGC2901C.P_SREFER3", Key:="P-R"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="AGC2901C.P_SREFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
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

    Dim i As Integer

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    Call Gp_Sp_Setting(sc3.Item("Spread"))
    Call Gp_Sp_Setting(sc4.Item("Spread"))
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc4.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    Call Gf_Sp_Cls(sc4)

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc3.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "G-System.INI", Me.Name)
     
    Call Gp_Sp_ColHidden(ss1, iss1Prod_rat, True)
    Call Gp_Sp_ColHidden(ss1, iss1Plan_rat, True)
    
    OPT_PLATE.Value = True
    opt_Product(1).Value = True

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc3.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "G-System.INI", Me.Name)
    
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
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set sc4 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(sc3)
        Call Gf_Sp_Cls(sc4)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Form_Exc()
    If SSTab1.Tab = 0 Then
        Call Gp_Sp_Excel(Me, sc1.Item("Spread"), 0, 0, 0, 0)
    ElseIf SSTab1.Tab = 1 Then
        Call Gp_Sp_Excel(Me, sc2.Item("Spread"), 0, 0, 0, 0)
    ElseIf SSTab1.Tab = 2 Then
        Call Gp_Sp_Excel(Me, sc3.Item("Spread"), 0, 0, 0, 0)
    Else
         Call Gp_Sp_Excel(Me, sc4.Item("Spread"), 0, 0, 0, 0)
    End If
End Sub

Public Sub Form_Ref()

    Dim sQuery      As String
    Dim dSlabwgt    As Double
    Dim dProdwgt    As Double
    Dim dOkwgt      As Double
    Dim iIdx        As Integer
    Dim iCol        As Integer
    
    Dim iCount      As Integer
    
    
On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

'    txt_Order.Text = Mid(txt_Order.Text, 2)
    
    Select Case SSTab1.Tab
    
           Case 0
            
                Call Display_ss1_Set
           
                sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", pControl)
                If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, 0, iSumCnt, iSumCol) Then
            '    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, , , False) Then
                    For iIdx = 1 To ss1.MaxRows
                        ss1.Row = iIdx
                        For iCol = ss1.MaxCols - iss1MaxCols + 1 To ss1.MaxCols
                            ss1.Col = iCol
                            If Val(ss1.Text & "") = 0 Then
                                ss1.Text = ""
                            End If
                        Next iCol
                    Next iIdx
                    
                    If ss1.MaxCols = iss1MaxCols Then
                       ss1.Col = 0:   ss1.Row = ss1.MaxRows:    ss1.Text = "合计"
                    End If
                    
                    iCol = ss1.MaxCols - iss1MaxCols
                    ss1.Row = ss1.MaxRows
                    ss1.Col = iCol + 1:             dSlabwgt = Val(Format(ss1.Text, "####.###") & "")
                    ss1.Col = iCol + 2:             dOkwgt = Val(Format(ss1.Text, "####.###") & "")
                    ss1.Col = iCol + 3:             dOkwgt = Val(Format(ss1.Text, "####.###") & "")
                    ss1.Col = iCol + iss1Cut_wgt:   dProdwgt = Val(Format(ss1.Text, "####.###") & "")
                    
                    ss1.Col = iCol + iss1Ok_rat:    If dProdwgt > 0 Then ss1.Text = Format(dOkwgt / dProdwgt * 100, "##0.0#")
                    ss1.Col = iCol + iss1Prod_rat:  If dSlabwgt > 0 Then ss1.Text = Format(dOkwgt / dSlabwgt * 100, "##0.0#")
                    'ss1.Col = iCol + 13:  If dSlabwgt > 0 Then ss1.Text = Format(dOkwgt / dProdwgt * 100, "##0.0#")

                    
                End If
                
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           
           Case 1
     
                If Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
            
           Case 2
           
                If TXT_SP_CD = "S" Then
                    If Gf_Sp_Refer(M_CN1, sc3, Mc1, , , False) Then
                        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                    End If
                End If
                
           Case 3
           
                If TXT_SP_CD <> "S" Or chk_Cond(5).Value = ssCBChecked Or _
                                       chk_Cond(6).Value = ssCBChecked Or _
                                       chk_Cond(7).Value = ssCBChecked Or _
                                       chk_Cond(10).Value = ssCBChecked Or _
                                       chk_Cond(11).Value = ssCBChecked Then
                        Call Gf_Sp_Cls(sc4)
                Else
                        Call Display_ss4_Set
                           
                        sQuery = Gf_Ms_MakeQuery(sc4.Item("P-R"), "R", pControl)
        '                If Gf_Total_Display(M_CN1, sc4, sQuery, 0, iSumCnt, iSumCol) Then
                        If Gf_Sp_Refer(M_CN1, sc4, Mc1, , , False) Then
                            For iIdx = 1 To ss4.MaxRows
                                ss4.Row = iIdx
                                For iCol = ss4.MaxCols - iss4MaxCols + 1 To ss4.MaxCols
                                    ss1.Col = iCol
                                    If Val(ss4.Text & "") = 0 Then
                                        ss4.Text = ""
                                    End If
                                Next iCol
                            Next iIdx
                            
        '                    If ss1.MaxCols = iss1MaxCols Then
        '                       ss1.Col = 0:   ss1.ROW = ss1.MaxRows:    ss1.Text = "合计"
        '                    End If
        '
        '                    iCol = ss1.MaxCols - iss1MaxCols
        '                    ss1.ROW = ss1.MaxRows
        '                    ss1.Col = iCol + 1:    dSlabwgt = Val(Format(ss1.Text, "####.###") & "")
        '                    ss1.Col = iCol + 2:    dOkwgt = Val(Format(ss1.Text, "####.###") & "")
        '                    ss1.Col = iCol + 3:    dOkwgt = Val(Format(ss1.Text, "####.###") & "")
        '                    ss1.Col = iCol + 14:   dProdwgt = Val(Format(ss1.Text, "####.###") & "")
        '
        '                    ss1.Col = iCol + 11:  If dProdwgt > 0 Then ss1.Text = Format(dOkwgt / dProdwgt * 100, "##0.0#")
        '                    ss1.Col = iCol + 12:  If dSlabwgt > 0 Then ss1.Text = Format(dOkwgt / dSlabwgt * 100, "##0.0#")
                            'ss1.Col = iCol + 13:  If dSlabwgt > 0 Then ss1.Text = Format(dOkwgt / dProdwgt * 100, "##0.0#")
        
                            
                        End If
                        
                        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

                End If
    
    End Select
    
    With ss2
        If .MaxRows <= 1 Then
            Exit Sub
        End If
            For iCount = 1 To .MaxRows
                 .Row = iCount
                 ss2.Row = .Row:       ss2.Col = SS2_URGNT_FL
               If ss2.Text = "Y" Then
                    Call Gp_Sp_BlockColor(ss2, SS2_SLAB_NO, SS2_SLAB_NO, .Row, .Row, &HC000&)
                    Call Gp_Sp_BlockColor(ss2, SS2_MAT_NO, SS2_MAT_NO, .Row, .Row, &HC000&)
                    Call Gp_Sp_BlockColor(ss2, SS2_ORD_NO, SS2_ORD_NO, .Row, .Row, &HC000&)
                    Call Gp_Sp_BlockColor(ss2, SS2_ORD_ITEM, SS2_ORD_ITEM, .Row, .Row, &HC000&)
                    Call Gp_Sp_BlockColor(ss2, SS2_URGNT_FL, SS2_URGNT_FL, .Row, .Row, &HC000&)
               End If
            Next iCount
    End With

    Exit Sub
    
Refer_Err:
    
End Sub

Private Sub Display_ss1_Set()
    Dim sSelCol     As String
    Dim iCol        As Integer
    Dim iIdx        As Integer
    Dim iInsCnt     As Integer
       
    ss1.DeleteCols 1, ss1.MaxCols - iss1MaxCols
    ss1.MaxCols = iss1MaxCols
    ss1.MaxRows = 0
    
    sSelCol = Trim(txt_Disp.Text)
    
    If sSelCol <> "" Then
        For iCol = 1 To Len(sSelCol) Step 2
            iInsCnt = iInsCnt + 1
            iIdx = Mid(sSelCol, iCol, 2)
            
            ss1.MaxCols = ss1.MaxCols + 1
            ss1.InsertCols ss1.MaxCols - iss1MaxCols, 1
            ss1.Col = ss1.MaxCols - iss1MaxCols
            ss1.Row = 0
            ss1.Text = chk_Cond(iIdx).Caption
        Next iCol
    End If
    
    Set iSumCol = Nothing
    
    iSumCnt = 11
    iSumCol.Add Item:=iInsCnt + 1
    iSumCol.Add Item:=iInsCnt + 2
    iSumCol.Add Item:=iInsCnt + 3
    iSumCol.Add Item:=iInsCnt + 4
    iSumCol.Add Item:=iInsCnt + 5
    iSumCol.Add Item:=iInsCnt + 6
    iSumCol.Add Item:=iInsCnt + 7
    iSumCol.Add Item:=iInsCnt + 8
    iSumCol.Add Item:=iInsCnt + 9
    iSumCol.Add Item:=iInsCnt + 10
    iSumCol.Add Item:=iInsCnt + 14
    
End Sub
Private Sub Display_ss4_Set()

    Dim sSelCol     As String
    Dim iCol        As Integer
    Dim iIdx        As Integer
    Dim iInsCnt     As Integer
       
    ss4.DeleteCols 1, ss4.MaxCols - iss4MaxCols
    ss4.MaxCols = iss4MaxCols
    ss4.MaxRows = 0
    
    sSelCol = Trim(txt_Disp.Text)
    
    If sSelCol <> "" Then
        For iCol = 1 To Len(sSelCol) Step 2
            iInsCnt = iInsCnt + 1
            iIdx = Mid(sSelCol, iCol, 2)
            
            ss4.MaxCols = ss4.MaxCols + 1
            ss4.InsertCols ss4.MaxCols - iss4MaxCols, 1
            ss4.Col = ss4.MaxCols - iss4MaxCols
            ss4.Row = 0
            ss4.Text = chk_Cond(iIdx).Caption
        Next iCol
    End If
    
    Set iSumCol = Nothing
    
'    iSumCnt = 2
'    iSumCol.Add Item:=iInsCnt + 1
'    iSumCol.Add Item:=iInsCnt + 2
'    iSumCol.Add Item:=iInsCnt + 3
'    iSumCol.Add Item:=iInsCnt + 4
'    iSumCol.Add Item:=iInsCnt + 5
'    iSumCol.Add Item:=iInsCnt + 6
'    iSumCol.Add Item:=iInsCnt + 7
'    iSumCol.Add Item:=iInsCnt + 8
'    iSumCol.Add Item:=iInsCnt + 9
'    iSumCol.Add Item:=iInsCnt + 10
'    iSumCol.Add Item:=iInsCnt + 14
    
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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_Product_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       TXT_PROD_CD = "AL"
       opt_Product(0).ForeColor = &HFF&
       opt_Product(1).ForeColor = &H808080
       opt_Product(2).ForeColor = &H808080
    ElseIf Index = 1 Then
       TXT_PROD_CD = "PP"
       opt_Product(1).ForeColor = &HFF&
       opt_Product(0).ForeColor = &H808080
       opt_Product(2).ForeColor = &H808080
    Else
       TXT_PROD_CD = "HC"
       opt_Product(2).ForeColor = &HFF&
       opt_Product(1).ForeColor = &H808080
       opt_Product(0).ForeColor = &H808080
    End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub chk_Cond_Click(Index As Integer)

    Dim Ord_Index As Integer

    If chk_Cond(Index) Then
        txt_Disp_Order = Trim(txt_Disp_Order & " " & chk_Cond(Index).Caption)
        txt_Order = Trim(txt_Order & chk_Cond(Index).Tag)
        txt_Disp = Trim(txt_Disp & Format(Index, "0#"))
    Else
        txt_Disp_Order = Trim(Replace(txt_Disp_Order, chk_Cond(Index).Caption, ""))
        txt_Order = Trim(Replace(txt_Order, chk_Cond(Index).Tag, ""))
        txt_Disp = Trim(Replace(txt_Disp, Format(Index, "0#"), ""))
    End If
    
    If Index = 12 Then
        Ord_Index = 14
        chk_Cond(Ord_Index) = chk_Cond(Index)
    End If
    
End Sub

Private Sub OPT_SLAB_Click(Value As Integer)

    If OPT_SLAB.Value = True Then
        OPT_SLAB.ForeColor = &HFF&
        OPT_PLATE.ForeColor = &H808080
        TXT_SP_CD = "S"
    Else
        OPT_SLAB.ForeColor = &H808080
        TXT_SP_CD = "P"
    End If

End Sub

Private Sub OPT_PLATE_Click(Value As Integer)

    If OPT_PLATE.Value = True Then
        OPT_PLATE.ForeColor = &HFF&
        OPT_SLAB.ForeColor = &H808080
        TXT_SP_CD = "P"
    Else
        OPT_PLATE.ForeColor = &H808080
        TXT_SP_CD = "S"
    End If

End Sub


Private Sub txt_Defect_name_Change()
    If txt_Defect_name.Text = "" Then
       txt_Defect.Text = ""
    End If
End Sub

Private Sub txt_Defect_name_DblClick()
    DD.sWitch = "MS"
    DD.sKey = "G0002"
    DD.rControl.Add Item:=txt_Defect

    DD.nameType = "2"

    Call Gf_Common_DD(M_CN1, vbKeyF4)
    
    If Len(txt_Defect.Text) = 3 Then
       txt_Defect_name.Text = Gf_ComnNameFind(M_CN1, "G0002", txt_Defect, 1)
    Else
       txt_Defect_name.Text = ""
    End If
    
End Sub

