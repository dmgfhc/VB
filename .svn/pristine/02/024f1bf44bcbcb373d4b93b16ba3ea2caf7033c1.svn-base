VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQE1062C 
   Caption         =   "南钢中厚板卷厂钢板/钢卷质量情况_AQE1062C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_PLT_NAME 
      Height          =   315
      Left            =   2640
      TabIndex        =   22
      Top             =   795
      Width           =   1335
   End
   Begin VB.TextBox TXT_PLT 
      Height          =   315
      Left            =   2250
      MaxLength       =   2
      TabIndex        =   21
      Top             =   795
      Width           =   405
   End
   Begin VB.TextBox Text_PROD_CD 
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
      Left            =   8115
      MaxLength       =   2
      TabIndex        =   14
      Tag             =   "产品"
      Top             =   285
      Width           =   495
   End
   Begin VB.TextBox Text_PROD_CD_Name 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8610
      TabIndex        =   13
      Top             =   285
      Width           =   1185
   End
   Begin VB.TextBox txt_STLGRD 
      Height          =   315
      Left            =   11520
      MaxLength       =   11
      TabIndex        =   12
      Top             =   285
      Width           =   1155
   End
   Begin VB.TextBox txt_STLGRD_Detail 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12660
      TabIndex        =   11
      Top             =   285
      Width           =   1530
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1140
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   1215
      Begin VB.OptionButton optJQ 
         BackColor       =   &H00E0E0E0&
         Caption         =   "剪切"
         Height          =   255
         Left            =   270
         TabIndex        =   4
         Top             =   765
         Width           =   765
      End
      Begin VB.OptionButton optZZ 
         BackColor       =   &H00E0E0E0&
         Caption         =   "轧制"
         Height          =   255
         Left            =   285
         TabIndex        =   3
         Top             =   465
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optZL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "装炉"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   165
         Width           =   765
      End
   End
   Begin VB.TextBox iMode 
      Height          =   270
      Left            =   10335
      TabIndex        =   0
      Text            =   "3"
      Top             =   1095
      Visible         =   0   'False
      Width           =   645
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   3915
      Left            =   75
      TabIndex        =   5
      Top             =   1290
      Width           =   15000
      _Version        =   393216
      _ExtentX        =   26458
      _ExtentY        =   6906
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   51
      MaxRows         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE1062C.frx":0000
   End
   Begin Threed.SSCommand cmd_Edit 
      Height          =   360
      Left            =   12510
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "重新生成数据"
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   3885
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   15000
      _Version        =   393216
      _ExtentX        =   26458
      _ExtentY        =   6853
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   52
      MaxRows         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE1062C.frx":189F
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   1665
      Top             =   285
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      Caption         =   "时间从"
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
   Begin InDate.UDate txt_from_date 
      Height          =   315
      Left            =   2460
      TabIndex        =   8
      Top             =   285
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   4260
      Top             =   285
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      Caption         =   "时间到"
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
   Begin InDate.UDate txt_to_date 
      Height          =   315
      Left            =   5055
      TabIndex        =   9
      Top             =   285
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   6810
      Top             =   285
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "产品"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   10170
      Top             =   285
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "钢种"
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
      ForeColor       =   16711680
   End
   Begin CSTextLibCtl.sidbEdit sdb_thk_fr 
      Height          =   315
      Left            =   5580
      TabIndex        =   15
      Top             =   795
      Width           =   930
      _Version        =   262145
      _ExtentX        =   1640
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
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   4260
      Top             =   795
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   7890
      Top             =   795
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "宽度"
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
   Begin CSTextLibCtl.sidbEdit sdb_wid_fr 
      Height          =   315
      Left            =   9210
      TabIndex        =   16
      Top             =   795
      Width           =   900
      _Version        =   262145
      _ExtentX        =   1587
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
      Modified        =   -1  'True
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
      NumIntDigits    =   4
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_thk_to 
      Height          =   315
      Left            =   6720
      TabIndex        =   17
      Top             =   795
      Width           =   900
      _Version        =   262145
      _ExtentX        =   1587
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
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_wid_to 
      Height          =   315
      Left            =   10350
      TabIndex        =   18
      Top             =   795
      Width           =   900
      _Version        =   262145
      _ExtentX        =   1587
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
      Modified        =   -1  'True
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
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   1665
      Top             =   795
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Caption         =   "工厂"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   225
      Left            =   10185
      TabIndex        =   20
      Top             =   885
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   225
      Left            =   6570
      TabIndex        =   19
      Top             =   885
      Width           =   90
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   15090
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "--"
      Height          =   180
      Left            =   4005
      TabIndex        =   10
      Top             =   420
      Width           =   255
   End
End
Attribute VB_Name = "AQE1062C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Plate & Coil Quality Analysis and Stat.
'-- Sub_System Name
'-- Program Name
'-- Program ID        AQE1060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2006.09.07
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
Dim nColumn1 As New Collection      'Spread necessary Column Collection
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCnt As Integer
Dim iSumCol As New Collection       'Sum Column

Dim SumCnt1 As Integer
Dim SumCol1 As New Collection       'Sum Column

Dim Cur_Spread As Object            'Spread Object
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    Dim i As Long
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(iMode, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_from_date, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_to_date, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
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
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    'Sc1.Add Item:="AQE1060C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQE1062C.P_SREFER1", Key:="P-R"
    'Sc1.Add Item:="AQE1060C.P_ONEROW", Key:="P-O"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
        'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 36, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 37, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 38, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 39, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 40, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 41, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 42, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 43, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 44, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 45, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 46, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 47, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 48, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 49, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 50, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 51, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 52, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    'Sc1.Add Item:="AQE1060C.P_MODIFY", Key:="P-M"
    sc2.Add Item:="AQE1062C.P_SREFER2", Key:="P-R"
    'Sc1.Add Item:="AQE1060C.P_ONEROW", Key:="P-O"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    SumCnt1 = 30
    
    'Sum Column Setting
    SumCol1.Add Item:=3
    
    For i = 7 To 12
        SumCol1.Add Item:=i
    Next i
    
    SumCol1.Add Item:=18
    SumCol1.Add Item:=20
    
    For i = 22 To 26
        SumCol1.Add Item:=i
    Next i
    
    SumCol1.Add Item:=28
    SumCol1.Add Item:=30
    
    For i = 32 To 34
        SumCol1.Add Item:=i
    Next i
    
    SumCol1.Add Item:=36
    SumCol1.Add Item:=38
    
    For i = 40 To 44
        SumCol1.Add Item:=i
    Next i
    
    SumCol1.Add Item:=46
    SumCol1.Add Item:=48
    SumCol1.Add Item:=50
    SumCol1.Add Item:=51
    
    'Sum Column Count
    iSumCnt = 30
    
    'Sum Column Setting
    iSumCol.Add Item:=4
    
    For i = 8 To 13
        iSumCol.Add Item:=i
    Next i
    
    iSumCol.Add Item:=19
    iSumCol.Add Item:=21
    
    For i = 23 To 27
        iSumCol.Add Item:=i
    Next i
    
    iSumCol.Add Item:=29
    iSumCol.Add Item:=31
    
    For i = 33 To 35
        iSumCol.Add Item:=i
    Next i
    
    iSumCol.Add Item:=37
    iSumCol.Add Item:=39
    
    For i = 41 To 45
        iSumCol.Add Item:=i
    Next i
    
    iSumCol.Add Item:=47
    iSumCol.Add Item:=49
    iSumCol.Add Item:=51
    iSumCol.Add Item:=52
 
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

'Private Sub cmd_Edit_Click()
'    Dim strDate As String
'    If txt_from_date.Text = txt_to_date.Text Then
'        strDate = txt_to_date.RawData
'        Call DataEdit(strDate, iMode.Text)
'    Else
'        Call Gp_MsgBoxDisplay("数据编辑量太大，请选择单独某一天编辑数据!!", "I")
'    End If
'End Sub

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
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)

    iMode.Text = "3"
    Screen.MousePointer = vbDefault
    Set Cur_Spread = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Set Cur_Spread = Nothing
   Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) And Gf_Sp_Cls(Proc_Sc("SC2")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        'rControl(1).SetFocus
''2007,06,04  SUNBIN start
         Text_PROD_CD_Name.Text = ""
         txt_PLT_NAME.Text = ""
''2007,0604   SUNBIN END
        Set Cur_Spread = Nothing
    End If

End Sub

Public Sub Form_Ref()
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
'    If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), 0, SumCnt1, SumCol1) Then
'       If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc2"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")), 1, 3, iSumCnt, iSumCol, False) Then
'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) And _
'       Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc1, Mc1("nControl"), Mc1("mControl")) Then

    If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), 1, 2, SumCnt1, SumCol1, False) Then
       If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc2"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")), 1, 1, iSumCnt, iSumCol, False) Then
        Call SS_DW(ss1, 1)
        Call SS_DW(ss1, 2)
        Call SS_DW(ss2, 1)
        Call Sp_Zero_Clear
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Sp_EvenRowBackcolor(Sc1.Item("Spread"))
        End If
    End If
    
    With ss1
    .Col = 2
    .Row = 1
    If .Text = "" Then
    Call Gf_Sp_Cls(Proc_Sc("SC2"))
    End If
    End With
    
            
End Sub




Public Sub Form_Pro()

'    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'    End If
    
End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    If Not Cur_Spread Is Nothing Then
        Call Gp_Sp_Excel(Me, Cur_Spread, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Else
        Call Gp_MsgBoxDisplay("请先点击选择你要导出的数据表！", "I")
    End If
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub optJQ_Click()
    iMode.Text = "1"
End Sub

Private Sub optZL_Click()
    iMode.Text = "2"
End Sub

Private Sub optZZ_Click()
    iMode.Text = "3"
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    Set Cur_Spread = ss1
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


'Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'
'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If
'
'End Sub

'Private Sub sp2_GroupReferSetting()
'    Dim i, j As Integer
'    If Trim(txt_from_date.RawData) = "" Or Trim(txt_to_date.RawData) = "" Then
'       MsgBox "查询日期未输入!", vbCritical, "系统提示信息"
'       Exit Sub
'    End If
'    Call Gf_Sp_Display(M_CN1, ss2, Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
'    Call Gp_Sp_EvenRowBackcolor(sc2.Item("Spread"))
'    With ss2
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 1
'            Select Case .Text
'            Case "A"
'               .Col = 0
'               .Text = "甲"
'            Case "B"
'               .Col = 0
'               .Text = "乙"
'            Case "C"
'               .Col = 0
'               .Text = "丙"
'            Case "D"
'               .Col = 0
'               .Text = "丁"
'            End Select
'
'            For j = 3 To .MaxCols
'                .Col = j
'            If (j <> 12 And j <> 14 And j <> 24 And j <> 26 And j <> 30 And j <> 32) _
'             And Val(.Text) = 0 Then
'                .Text = ""
'             End If
'            Next j
'        Next i
'    End With
'
'End Sub

'Private Sub DataEdit(ByVal strDate As String, ByVal iMode As String)
'    On Error GoTo Process_Exec_ERROR
'
'    Dim OutParam(1, 4) As Variant
'    Dim ret_Result_ErrMsg As String
'    Dim sQuery As String
'    Dim icount As Integer
'
'    Dim adoCmd As adodb.Command
'
'    'If ss1.MaxRows = 0 Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'    'Return Error Messsage Parameter
'    OutParam(1, 1) = "arg_e_msg"
'    OutParam(1, 2) = adVarChar
'    OutParam(1, 3) = adParamOutput
'    OutParam(1, 4) = 256
'
'    sQuery = "{call AQE1060P ('" + strDate + "','" + iMode + "',?)}"
'
'    'Ado Setting
'    M_CN1.CursorLocation = adUseServer
'    Set adoCmd = New adodb.Command
'
'    adoCmd.CommandType = adCmdText
'    Set adoCmd.ActiveConnection = M_CN1
'
'    adoCmd.CommandText = sQuery
'
'    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
'
'    adoCmd.Execute , , adExecuteNoRecords
'
'    'Process Error Check
'    If adoCmd("arg_e_msg") <> "" Then
'        ret_Result_ErrMsg = adoCmd("arg_e_msg")
'        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
'        Call Gp_MsgBoxDisplay(sErrMessg)
'    Else
'        Call Gp_MsgBoxDisplay("数据编辑结束!!", "I")
''        txt_prod_cd.Text = "PP"
''        Call txt_prod_cd_KeyUp(0, 0)
'        Call Form_Ref
'    End If
'
'    Set adoCmd = Nothing
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'Process_Exec_ERROR:
'
'    Set adoCmd = Nothing
'    Screen.MousePointer = vbDefault
'    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
'
'End Sub
Private Sub Sp_Zero_Clear()
    Dim i As Integer
    Dim j As Integer
    With ss1
     For i = 1 To .MaxRows
         .Row = i
         For j = 2 To .MaxCols
             .Col = j
             If .CellType = CellTypeNumber And _
             Val(.Text) = 0 Then
                .Text = ""
             End If
         Next j
     Next i
    End With
    With ss2
         For i = 1 To .MaxRows
             .Row = i
             For j = 2 To .MaxCols
                 .Col = j
                 If .CellType = CellTypeNumber And _
                 Val(.Text) = 0 Then
                    .Text = ""
                 End If
             Next j
         Next i
        End With
End Sub


Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
        If ss1.MaxRows < 1 Or Row = 0 Or Row = -999 Or ss1.MaxRows = Row Then Exit Sub
    
        Unload AQE1061C
        Load AQE1061C
               
        ss1.Row = Row
        ss1.Col = 1
        AQE1061C.Text_PROD_CD.Text = IIf(Trim(ss1.Text) = "钢板", "PP", "HC")
        
        ss1.Row = Row
        ss1.Col = 2
        AQE1061C.txt_STLGRD.Text = STLGRD_Find(Trim(ss1.Text))
        AQE1061C.txt_STLGRD_DETAIL.Text = Trim(ss1.Text)
        
        ss1.Row = Row
        ss1.Col = 5
        AQE1061C.sdb_thk_fr = ss1.Value: AQE1061C.sdb_thk_to = ss1.Value
        
        ss1.Row = Row
        ss1.Col = 6
        AQE1061C.sdb_wid_fr = ss1.Value: AQE1061C.sdb_wid_to = ss1.Value
        
        AQE1061C.PROD_DATE_FR.RawData = txt_from_date.RawData
        AQE1061C.PROD_DATE_TO.RawData = txt_to_date.RawData
        
        AQE1061C.optZL.Value = Me.optZL.Value
        AQE1061C.optZL.Value = Me.optZL.Value
        AQE1061C.optZZ.Value = Me.optZZ.Value
        AQE1061C.iMode = Me.iMode
        AQE1061C.Active_CForm = "AQE1061C"
        AQE1061C.Show
        AQE1061C.SetFocus

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2 '
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    Set Cur_Spread = ss2
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0 '
End Sub
Private Function STLGRD_Find(ByVal stlgrdName As String) As String
    Dim sQuery As String
  
    sQuery = "SELECT STLGRD FROM QP_NISCO_CHMC WHERE  STEEL_GRD_DETAIL= '" + Trim(stlgrdName) + "'"
       
    STLGRD_Find = Gf_FloatFind(M_CN1, sQuery)
   
End Function

Private Sub Text_PROD_CD_Change()
   
    Select Case Text_PROD_CD.Text
        Case "S", "s", "SL"
            Text_PROD_CD.Text = "SL"
        Case "P", "p", "PP"
            Text_PROD_CD.Text = "PP"
            
        Case "H", "h", "HC"
            Text_PROD_CD.Text = "HC"
            
        Case ""
            Text_PROD_CD.Text = ""
        Case Else
            Text_PROD_CD.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select

End Sub



Private Sub Text_PROD_CD_LostFocus()

    If Text_PROD_CD.Text <> "" Then
        If (Len(Text_PROD_CD.Text) < Text_PROD_CD.MaxLength) Then
            Call Gp_MsgBoxDisplay("产品分类代码输入未完成！")
            'Text_PROD_CD.Text = ""
            Text_PROD_CD.SetFocus
        End If
    End If

End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)

   Text_PROD_CD_Name.Text = ""
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=Text_PROD_CD
        DD.rControl.Add Item:=Text_PROD_CD_Name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(Text_PROD_CD.Text)) = Text_PROD_CD.MaxLength Then
        Text_PROD_CD_Name.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_PROD_CD.Text, 2)
    Else
        Text_PROD_CD_Name.Text = ""
    End If
    
End Sub

Private Sub TXT_PLT_Change()
   
   Select Case TXT_PLT.Text
        Case "C1", "c1"
            TXT_PLT.Text = "C1"
        Case "C3", "c3"
            TXT_PLT.Text = "C3"
        Case "B1", "b1"
            TXT_PLT.Text = "B1"
        Case "*", "**"
            TXT_PLT.Text = "**"
        Case ""
            TXT_PLT.Text = ""
            txt_PLT_NAME.Text = ""
'        Case Else
'            TXT_PLT.Text = ""
'            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select
End Sub
Private Sub TXT_PLT_KeyUp(KeyCode As Integer, Shift As Integer)

   txt_PLT_NAME.Text = ""
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0001"

        DD.rControl.Add Item:=TXT_PLT
        DD.rControl.Add Item:=txt_PLT_NAME
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(TXT_PLT.Text)) = TXT_PLT.MaxLength Then
        txt_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", TXT_PLT.Text, 2)
    Else
        txt_PLT_NAME.Text = ""
    End If
    
End Sub

Private Sub txt_STLGRD_Change()
Dim sQuery As String
    If Len(Trim(txt_STLGRD.Text)) = 11 Then
       sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(txt_STLGRD.Text) + "'"
       
       txt_STLGRD_DETAIL = Gf_FloatFind(M_CN1, sQuery)
    Else
        txt_STLGRD_DETAIL = ""
    End If
End Sub

Private Sub txt_STLGRD_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim sQuery As String
    If KeyCode = vbKeyF4 Then
    
        DD.nameType = "1"
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_STLGRD
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_STLGRD.Text)) = 11 Then
           sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(txt_STLGRD.Text) + "'"
           txt_STLGRD_DETAIL = Gf_FloatFind(M_CN1, sQuery)
        Else
            txt_STLGRD_DETAIL = ""
        End If
    End If
End Sub

Private Sub sdb_thk_fr_Change()
    If sdb_thk_fr.Value > 0 And sdb_thk_to.Value < sdb_thk_fr.Value Then
        sdb_thk_to.Value = sdb_thk_fr.Value
    End If
End Sub

Private Sub sdb_wid_fr_Change()
    If sdb_wid_fr.Value > 0 And sdb_wid_to.Value < sdb_wid_fr.Value Then
        sdb_wid_to.Value = sdb_wid_fr.Value
    End If
End Sub
Private Sub SS_SUM_PERCENT(ByVal vaMySpread As vaSpread, Col_Slab_PROC As Integer, Col_Total_PROC As Integer, Col_Acture_PROC As Integer, Col_Certif_PROC As Integer, Col_NotPlan_PROC As Integer, Col_Dsc_Scrap As Integer, Col_Mid_Scrap As Integer, Col_Succes_PERC As Integer, Col_Once_Certif_PERC As Integer, Col_ShouD_PERC As Integer, Col_ZhaCh_PERC As Integer, Row As Long)

Dim Vaul_Col_Slab_PROC         As Double                 ' 坯料重量
Dim Vaul_Col_Total_PROC        As Double                 ' 生产总量
Dim Vaul_Col_Acture_PROC       As Double                 ' 实际产量
Dim Vaul_Col_Certif_PROC       As Double                 ' 合格量
Dim Vaul_Col_NotPlan_PROC      As Double                 ' 非计划量
Dim Vaul_Col_Dsc_Scrap         As Double                 ' 判废
Dim Vaul_Col_Mid_Scrap         As Double                 ' 中废



  With vaMySpread
     .Row = Row
     .Col = Col_Slab_PROC
     If .Text > 0 Then
       Vaul_Col_Slab_PROC = .Text
       Else: Vaul_Col_Slab_PROC = 0
       End If
    
    .Col = Col_Total_PROC
    If .Text > 0 Then
       Vaul_Col_Total_PROC = .Text
       Else: Vaul_Col_Total_PROC = 0
       End If
       
    .Col = Col_Acture_PROC
    If .Text > 0 Then
       Vaul_Col_Acture_PROC = .Text
       Else:
       Vaul_Col_Acture_PROC = 0
'       Call MsgBox("有一行实际产量数据不准确", vbExclamation + vbOKOnly, "警告")
       Exit Sub
       End If
       
    .Col = Col_Certif_PROC
    If .Text > 0 Then
       Vaul_Col_Certif_PROC = .Text
       Else: Vaul_Col_Certif_PROC = 0
       End If
       
     .Col = Col_NotPlan_PROC
    If .Text > 0 Then
       Vaul_Col_NotPlan_PROC = .Text
       Else: Vaul_Col_NotPlan_PROC = 0
       End If
       
    .Col = Col_Dsc_Scrap
    If .Text > 0 Then
       Vaul_Col_Dsc_Scrap = .Text
       Else: Vaul_Col_Dsc_Scrap = 0
       End If
       
    .Col = Col_Mid_Scrap
    If .Text > 0 Then
       Vaul_Col_Mid_Scrap = .Text
       Else: Vaul_Col_Mid_Scrap = 0
       End If
       
    .Col = Col_Succes_PERC
    .Text = Vaul_Col_Acture_PROC / Vaul_Col_Slab_PROC * 100                         '成材率
    .Col = Col_Once_Certif_PERC
    .Text = (Vaul_Col_Acture_PROC - Vaul_Col_Dsc_Scrap - Vaul_Col_Mid_Scrap) / Vaul_Col_Acture_PROC * 100  '一次合格率
    .Col = Col_ShouD_PERC
    .Text = Vaul_Col_Certif_PROC / Vaul_Col_Total_PROC * 100                        '收得率
    .Col = Col_ZhaCh_PERC
    .Text = (Vaul_Col_Acture_PROC - Vaul_Col_NotPlan_PROC - Vaul_Col_Dsc_Scrap - Vaul_Col_Mid_Scrap) / Vaul_Col_Slab_PROC * 100   '轧成率
       
  End With
 
End Sub


Private Sub SS_DW(ByVal vaMySpread As vaSpread, Col_Choice As Long)
Dim iRow As Long
Dim iSaveRowNo() As Long
Dim iArrayNum As Integer

With vaMySpread
    .Col = Col_Choice
    ReDim iSaveRowNo(.MaxRows)
    For iRow = 1 To .MaxRows
        .Row = iRow
        If InStr(1, .Text, "小计") > 0 Then
            iSaveRowNo(iRow - 1) = iRow
        ElseIf InStr(1, .Text, "合计") > 0 Then
            iSaveRowNo(iRow - 1) = iRow
        Else
            iSaveRowNo(iRow - 1) = 0
        End If
        
    Next iRow

    For iArrayNum = 0 To .MaxRows
        If iSaveRowNo(iArrayNum) > 0 Then
            If .Name = "ss1" Then
                Call SS_SUM_PERCENT(vaMySpread, 3, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, iSaveRowNo(iArrayNum))
            ElseIf .Name = "ss2" Then
                Call SS_SUM_PERCENT(vaMySpread, 4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, iSaveRowNo(iArrayNum))
            Else
                Exit Sub
            End If
        End If
    Next iArrayNum
End With

End Sub




