VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQC0070C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "产品信息查询及修改界面_AQC0070C"
   ClientHeight    =   8730
   ClientLeft      =   360
   ClientTop       =   1890
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_PLT 
      Alignment       =   2  'Center
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
      Left            =   1620
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "plt"
      Top             =   90
      Width           =   465
   End
   Begin VB.CheckBox Check_NULL 
      Height          =   285
      Left            =   4260
      TabIndex        =   18
      Top             =   1110
      Width           =   195
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_LEN_CHAG 
      Height          =   315
      Left            =   3915
      TabIndex        =   17
      Top             =   1530
      Width           =   1185
      _Version        =   262145
      _ExtentX        =   2090
      _ExtentY        =   556
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
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
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
      NumIntDigits    =   5
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_WID_CHAG 
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   1530
      Width           =   1005
      _Version        =   262145
      _ExtentX        =   1773
      _ExtentY        =   556
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
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_THK_CHAG 
      Height          =   315
      Left            =   1965
      TabIndex        =   15
      Top             =   1530
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
      _ExtentY        =   556
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
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
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
      NumIntDigits    =   2
      Undo            =   0
      Data            =   0
   End
   Begin VB.TextBox txt_MAT_NO 
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
      Left            =   10260
      MaxLength       =   14
      TabIndex        =   4
      Top             =   90
      Width           =   1665
   End
   Begin VB.TextBox txt_CHG_STDSPEC 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
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
      Left            =   1980
      TabIndex        =   11
      Tag             =   "改判时适用标准"
      Top             =   1080
      Width           =   1950
   End
   Begin VB.TextBox txt_PROD_CD 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   3420
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "产品"
      Text            =   "PP"
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox txt_STDSPEC 
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
      Left            =   1980
      MaxLength       =   18
      TabIndex        =   6
      Top             =   525
      Width           =   1935
   End
   Begin VB.TextBox txt_CUST_SPEC 
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
      Left            =   5820
      MaxLength       =   14
      TabIndex        =   7
      Top             =   525
      Width           =   1935
   End
   Begin VB.TextBox txt_PROD_GRD 
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
      Left            =   13680
      MaxLength       =   14
      TabIndex        =   5
      Top             =   90
      Width           =   315
   End
   Begin VB.TextBox TXT_ORD_GRD_NAME 
      Height          =   315
      Left            =   13995
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   90
      Width           =   1155
   End
   Begin VB.TextBox txt_THK_MIN 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   10290
      MaxLength       =   5
      TabIndex        =   8
      Top             =   540
      Width           =   810
   End
   Begin VB.TextBox txt_THK_MAX 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   11475
      MaxLength       =   5
      TabIndex        =   9
      Top             =   525
      Width           =   855
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   7155
      Left            =   135
      TabIndex        =   10
      Top             =   2025
      Width           =   15030
      _Version        =   393216
      _ExtentX        =   26511
      _ExtentY        =   12621
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   22
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0070C.frx":0000
   End
   Begin Threed.SSCommand SSC_CHANG_SELECT 
      Height          =   315
      Left            =   10860
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "修改所有选定的产品"
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   8700
      Top             =   90
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "查询产品号"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   4230
      Tag             =   "生产日"
      Top             =   90
      Width           =   1560
      _ExtentX        =   2752
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
   Begin InDate.UDate SDT_PROD_DATE_FR 
      Height          =   315
      Left            =   5850
      TabIndex        =   2
      Tag             =   "生产日"
      Top             =   90
      Width           =   1410
      _ExtentX        =   2487
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
      MaxLength       =   10
   End
   Begin InDate.UDate SDT_PROD_DATE_TO 
      Height          =   315
      Left            =   7245
      TabIndex        =   3
      Tag             =   "生产日"
      Top             =   90
      Width           =   1410
      _ExtentX        =   2487
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
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   405
      Top             =   1080
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "修改产品标准"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   405
      Top             =   525
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "产品标准"
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
      Left            =   4230
      Top             =   525
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "客户特殊要求"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   4245
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Caption         =   "删除客户特殊要求"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   8700
      Top             =   525
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   556
      Caption         =   "产品厚度"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   11115
      Top             =   540
      Width           =   340
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   "-"
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
      ForeColor       =   0
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   315
      Left            =   13110
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消所有选定的修改"
   End
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   12105
      Top             =   90
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "产品等级"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   390
      Top             =   1500
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "修改产品尺寸"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   2235
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Left            =   405
      Top             =   90
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      Caption         =   "工厂"
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
   Begin VB.Line Line3 
      X1              =   0
      X2              =   15120
      Y1              =   975
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   -15
      X2              =   15235
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -15
      X2              =   15220
      Y1              =   1890
      Y2              =   1890
   End
End
Attribute VB_Name = "AQC0070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        AQC0070C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2005.8.30
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

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

Dim sWgtLenFlag As String
Dim sQuery      As String
Dim bCheck      As Boolean

Const SS1_PLATE_NO = 1                  'PLATE NO
Const SS1_PROC_CD = 3                   'PRODUCT STATUS
Const SS1_THK = 6
Const SS1_WID = 7
Const SS1_LEN = 8
Const SS1_WGT = 9
Const SS1_PROD_GRD = 2
Const SS1_BEF_STDSPEC = 12              'BEFORE 标准号
Const SS1_STDSPEC = 13                  '标准号
Const SS1_USER_ID = 22                  'USER ID
Const SS1_CUSTSPEC = 14
Const SS1_APLY_CUSTSPEC = 15

Private Sub Form_Define()
      
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_PROD_CD, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(SDT_PROD_DATE_FR, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
            Call Gp_Ms_Collection(txt_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
           Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(txt_CUST_SPEC, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_PROD_GRD, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
           Call Gp_Ms_Collection(txt_THK_MIN, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
           Call Gp_Ms_Collection(txt_THK_MAX, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
               Call Gp_Ms_Collection(txt_PLT, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection1(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0070C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQC0070C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQC0070C.P_ROWONE", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="sc1"
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "◎"
    
'    Call Gp_Sp_ColHidden(ss1, SS1_USER_ID, True)
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_PROD_CD"             '产品
            sCode = "B0005"
        Case "txt_PROD_GRD"
            sCode = "Q0034"
            Set oCodeName = TXT_ORD_GRD_NAME
        Case "txt_STDSPEC"
            sCode = "STDSPEC"
        Case "txt_CUST_SPEC"
            sCode = "CUST_SPEC_NO"
        
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuToolSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    
    With ss1
        .Row = 0: .Row2 = 0
        .Col = SS1_STDSPEC: .Col2 = SS1_STDSPEC
        
        .BlockMode = True
        
        .CellType = SS_CELL_TYPE_STATIC_TEXT
        .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .TypeTextWordWrap = True
        
        .BackColor = &HE1E4CD
        .ForeColor = BLUE
        
        .BlockMode = False

    End With
    
    Call Gf_Sp_Cls(Sc1)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "G-System.INI", Me.Name)

    SDT_PROD_DATE_FR.RawData = Gf_DTSet(M_CN1, "D")
    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
    txt_PROD_CD.Text = "PP"
    
    bCheck = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "G-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
        
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
     
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("sc1"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Sc1) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuToolSet
        txt_CHG_STDSPEC.Text = ""
        
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call pContro1(1).SetFocus
    End If
    bCheck = False
End Sub

Public Sub MenuToolSet()

    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
'    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
'    MDIMain.MenuTool.Buttons(14).Enabled = False

End Sub

Public Sub Form_Ref()
    
    If Gf_Sp_ProceExist(Sc1.Item("Spread")) Then Exit Sub
    
    If Trim(txt_MAT_NO.Text) = "" And SDT_PROD_DATE_FR.RawData = "" Then
        Call Gp_MsgBoxDisplay("请输入查询号还是生产日期！！！")
        Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuToolSet
        txt_CHG_STDSPEC.Text = ""
        bCheck = False
    End If
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub


Public Sub Form_Pro()
    Dim iDR         As Long
    Dim sSpec       As String
    Dim sBefSpec    As String

    If Gf_Sp_Process(M_CN1, Proc_Sc("sc1"), Mc1) Then
'''        txt_MAT_NO.Text = ""
        Call Form_Ref
        Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        Call MenuToolSet
    End If
    
End Sub

Public Sub Form_Exc()
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End Sub




Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    Call SET_CHANG_RUSLT(Row, 1)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Call SET_CHANG_RUSLT(Row, Mode)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Call ss1_KeyUp(vbKeyF4, ss1.ActiveRow)
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyF4 Then
                
        Set DD.sPname = Me.ss1
        
        DD.sWitch = "SP"
        DD.rControl.Add Item:=SS1_STDSPEC
        
        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
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

Private Sub txt_PROD_CD_Change()
    Select Case txt_PROD_CD.Text
        Case "S", "s", "SL"
            txt_PROD_CD.Text = "SL"
        Case "P", "p", "PP"
            txt_PROD_CD.Text = "PP"
        Case "H", "h", "HC"
            txt_PROD_CD.Text = "HC"
        Case ""
            txt_PROD_CD.Text = ""
        Case Else
            txt_PROD_CD.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select

End Sub


Private Sub SSC_CHANG_SELECT_Click()
    
'    Dim I As Integer
'
'    If ss1.SelBlockCol = 0 And ss1.SelBlockCol2 = 0 Then Exit Sub
'
'        With ss1
'            .Row = 1
'
'        End With
'
End Sub

Private Sub txt_CHG_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_CHG_STDSPEC

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
    End If
End Sub
Public Sub Gp_Sp_Collection1(sPname As Variant, Num As Integer, pcol As String, ncol As String, mcol As String, _
                                                               iCol As String, acol As String, lCol As String, _
                            pColumn As Collection, nColumn As Collection, mColumn As Collection, iColumn As Collection, _
                            aColumn As Collection, lColumn As Collection)
   
    If LCase(Trim(pcol)) = "p" Then       'PK Column
        pColumn.Add Item:=Num
    End If
    
    If LCase(Trim(ncol)) = "n" Then       'Necessary Column
        nColumn.Add Item:=Num
        'Call Gp_Sp_ColColor(SpName, Num, , &H80FF80)
    End If
    
    If LCase(Trim(mcol)) = "m" Then       'Spread Maxlength check Column
        mColumn.Add Item:=Num
    End If
    
    If LCase(Trim(iCol)) = "i" Then       'Spread Insert Column
        iColumn.Add Item:=Num
        Call Gp_Sp_ColColor(sPname, Num, , &HC0FFFF)
    End If
    
    If LCase(Trim(acol)) = "a" Then       'Master -> Spread Column
        aColumn.Add Item:=Num
        Call Gp_Sp_ColHidden(sPname, Num, True)
    End If
    
    If LCase(Trim(lCol)) = "l" Then       'Spread Lock Column
        lColumn.Add Item:=Num
        Call Gp_Sp_ColLock(sPname, Num, True)
        Call Gp_Sp_ColColor(sPname, Num, , &H80000005)
    End If

    
End Sub

Private Sub SET_CHANG_RUSLT(ByVal Row As Long, ByVal Mode As Integer)
    
    Dim sStdspec     As String
    Dim sAplyStd     As String
    Dim sCustSPEC    As String
    Dim sAPLYCUSTSTD As String
    Dim dTHK, dWID, dLEN As Double
    Dim dAPLYTHK, dAPLYWID, dAPLYLEN As Double
    Dim iC1, iC2, iC3, iC4, iC5, iC6 As Integer
    
    iC1 = 0
    iC2 = 0
    iC3 = 0
    iC4 = 0
    iC5 = 0
    iC6 = Mode
    
    sAplyStd = Trim(txt_CHG_STDSPEC.Text)
    'sAPLYCUSTSTD = Trim(txt_APLY_CUST_SPEC.Text)
    dAPLYTHK = sidbEdit_THK_CHAG.Value
    dAPLYWID = sidbEdit_WID_CHAG.Value
    dAPLYLEN = sidbEdit_LEN_CHAG.Value
    
    If Row > 0 Then
        With ss1
        
        .Row = Row
        .Col = SS1_BEF_STDSPEC
        sStdspec = .Text
        '修改标准
        .Col = SS1_STDSPEC
            If Trim(sAplyStd) <> "" And sAplyStd <> sStdspec Then
                .Text = sAplyStd
                iC1 = 1
            Else
                .Text = ""
                iC1 = 0
            End If
        '删除客户特殊要求
        .Col = SS1_CUSTSPEC
        sCustSPEC = .Text
        .Col = SS1_APLY_CUSTSPEC
            If Check_NULL.Value <> 0 Then
                .Text = "Y"
                iC2 = 1
            Else
                .Text = "N"
                iC2 = 0
            End If
        '修该产品尺寸
        .Col = SS1_THK
        dTHK = .Value
        
            If dAPLYTHK > 0 And dAPLYTHK <> dTHK Then
                .Text = dAPLYTHK
                iC3 = 1
            Else
                .Text = dTHK
                iC3 = 0
            End If
        
        .Col = SS1_WID
        dWID = .Value
        
            If dAPLYWID > 0 And dAPLYWID <> dWID Then
                .Text = dAPLYWID
                iC4 = 1
            Else
                .Text = dWID
                iC4 = 0
            End If
        
        .Col = SS1_LEN
        dLEN = .Value
        
            If dAPLYLEN > 0 And dAPLYLEN <> dLEN Then
                .Text = dAPLYLEN
                iC5 = 1
            Else
                .Text = dLEN
                iC5 = 0
            End If
        
        '设置修改标记
        
        .Col = 0
            If iC1 > 0 Or iC2 > 0 Or iC3 > 0 Or iC4 > 0 Or iC5 > 0 Then
                iC6 = 0
            End If
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc1")("Spread"), iC6)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), SS1_USER_ID)
            
        End With
    End If
        
    
End Sub

Private Sub TXT_PLT_Change()

    If txt_PLT.Text = "C3" Then
       txt_PROD_CD.Text = "PP"
    End If

End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_PLT

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub



