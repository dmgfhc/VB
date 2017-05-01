VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AEB1080C 
   Caption         =   "HMI 炉次编制_AEB1080C"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   1455
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15135
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_mlt_proc_cd 
      Height          =   330
      Left            =   135
      TabIndex        =   9
      Top             =   5310
      Visible         =   0   'False
      Width           =   465
   End
   Begin InDate.ULabel ULabel2 
      Height          =   360
      Left            =   3060
      Top             =   5235
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   635
      Caption         =   ""
      Alignment       =   1
      BackColor       =   14737632
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
   Begin VB.TextBox txt_line 
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
      Left            =   6030
      MaxLength       =   1
      TabIndex        =   5
      Tag             =   "机号"
      Top             =   90
      Width           =   420
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   1410
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "工厂"
      Top             =   90
      Width           =   3615
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   945
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "工厂"
      Top             =   90
      Width           =   465
   End
   Begin VB.TextBox txt_prod_cd_name 
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
      Left            =   7935
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "产品"
      Top             =   90
      Width           =   2130
   End
   Begin VB.TextBox txt_prod_cd 
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
      Left            =   7470
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "产品"
      Top             =   90
      Width           =   465
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4785
      Left            =   135
      TabIndex        =   0
      Top             =   450
      Width           =   15090
      _Version        =   393216
      _ExtentX        =   26617
      _ExtentY        =   8440
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
      MaxCols         =   17
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AEB1080C.frx":0000
      UserResize      =   1
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   135
      Top             =   90
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   5220
      Top             =   90
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      Caption         =   "机号"
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
      Left            =   6660
      Top             =   90
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin InDate.ULabel lbl_slab 
      Height          =   120
      Index           =   0
      Left            =   3330
      Top             =   8820
      Visible         =   0   'False
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   212
      Caption         =   ""
      Alignment       =   1
      BackColor       =   8421631
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel3 
      Height          =   510
      Left            =   3240
      Top             =   8910
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   900
      Caption         =   ""
      Alignment       =   1
      BackColor       =   14737632
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   8325
      Top             =   5805
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "合计重量"
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
   Begin CSTextLibCtl.sidbEdit sdb_charge_wgt 
      Height          =   315
      Left            =   9405
      TabIndex        =   12
      Top             =   5805
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      Modified        =   -1  'True
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
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_slab_init 
      Height          =   600
      Left            =   13500
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6705
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "初始化"
   End
   Begin Threed.SSCommand cmd_slab_complete 
      Height          =   600
      Left            =   13500
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8055
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   12583104
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
   End
   Begin Threed.SSCommand cmd_slab_del 
      Height          =   600
      Left            =   13500
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7380
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1058
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   32896
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "删除"
   End
   Begin CSTextLibCtl.sidbEdit sdb_heat_edt_seq 
      Height          =   315
      Left            =   135
      TabIndex        =   16
      Top             =   5715
      Visible         =   0   'False
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      NumIntDigits    =   8
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin VB.Label lbl_min 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2115
      TabIndex        =   11
      Top             =   7065
      Width           =   960
   End
   Begin VB.Label lbl_max 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2115
      TabIndex        =   10
      Top             =   5985
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3330
      X2              =   7965
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "最小(T)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2115
      TabIndex        =   8
      Top             =   6840
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "最大(T)"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2115
      TabIndex        =   7
      Top             =   5760
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "20(M)"
      Height          =   195
      Left            =   8190
      TabIndex        =   6
      Top             =   8685
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      FillColor       =   &H000000FF&
      Height          =   3840
      Left            =   3195
      Shape           =   4  'Rounded Rectangle
      Top             =   5175
      Width           =   4920
   End
End
Attribute VB_Name = "AEB1080C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEB1080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.9.22
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSlab_cnt As Integer            'Slab Design Count
Dim iHeat_edt_seq As Long           'Heat_edt_seq Value
Dim lMain_row As Long               'Main Row

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_line, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_prod_cd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_MLT_PROC_CD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(sdb_charge_wgt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(sdb_heat_edt_seq, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(SS1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=SS1, Key:="Spread"
    Sc1.Add Item:="AEB1080C.P_REFER", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=SS1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cmd_slab_complete_Click()

    On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer
    Dim iCnt As Long
    
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput  'adParamInput, adParamOutput, adParamInputOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEB1080P (" & sdb_heat_edt_seq.Value & ",?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        
        For iCnt = 1 To iSlab_cnt
            Unload lbl_slab(iCnt)
        Next iCnt
        
        iSlab_cnt = 0
        lMain_row = 0
        sdb_charge_wgt.Value = 0
        sdb_heat_edt_seq.Value = 0
        txt_MLT_PROC_CD.Text = ""
        lbl_max.Caption = ""
        lbl_min.Caption = ""
                
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_slab_del_Click()

    Dim sSeq As String
    
    Dim iCount As Integer
    Dim iRow As Integer
    Dim iVisible_Cnt As Integer
    
    If iSlab_cnt = 0 Then Exit Sub
    
    For iCount = 1 To iSlab_cnt
        
        If lbl_slab(iCount).Caption = "删除" Then
            
            If lbl_slab(iCount).Visible Then
            
                lbl_slab(iCount).Height = 0
                lbl_slab(iCount).Visible = False
                
                '--------------------------------------------------
                If iCount < 10 Then
                    sSeq = "0" & iCount
                Else
                    sSeq = str(iCount)
                End If
                    
            End If
            
            For iRow = 1 To SS1.MaxRows
                
                SS1.Row = iRow
                SS1.Col = 3
                
                If sSeq = SS1.Text Then
                    SS1.Text = ""
                    SS1.Col = 2
                    SS1.Value = 0
                    SS1.Col = 0
                    SS1.Text = ""
                    
                    'SLAB_WGT
                    SS1.Col = 8
                    sdb_charge_wgt.Value = sdb_charge_wgt.Value - SS1.Value
                    
                    'EP_SLAB_EDT  TABLE HEAT_EDT_SEQ, HEAT_SLSB_SEQ UPDATEING
                    SS1.Col = 1 'SLAB_EDT_SEQ
                    Call Slab_Seq_Create("D", SS1.Text, sSeq)
        
                    Call Gp_Sp_BlockColor(SS1, 1, SS1.MaxCols, iRow, iRow)
                End If

            Next iRow
            
        End If
    
        If iCount = 1 Then
            lbl_slab(iCount).Top = 8595 - lbl_slab(iCount).Height
        Else
            If lbl_slab(iCount - 1).Caption <> "删除" Then
                lbl_slab(iCount).Top = lbl_slab(iCount - 1).Top - lbl_slab(iCount).Height
            Else
                lbl_slab(iCount).Top = lbl_slab(iCount - 1).Top - lbl_slab(iCount).Height + 30
            End If
        End If
    
    Next iCount
    
    iVisible_Cnt = 0
    For iCount = 1 To iSlab_cnt
    
        If lbl_slab(iCount).Visible Then
            iVisible_Cnt = iVisible_Cnt + 1
        End If
    
    Next iCount
    
    'EP_SLAB_EDT DATA DELETE
    If iVisible_Cnt = 0 Then
    
        For iCount = 1 To iSlab_cnt
            Unload lbl_slab(iCount)
        Next iCount
        
        iSlab_cnt = 0
        lMain_row = 0
        lbl_max.Caption = ""
        lbl_min.Caption = ""
        sdb_charge_wgt.Value = 0
        cmd_slab_init.Enabled = False
        cmd_slab_del.Enabled = False
        cmd_slab_complete.Enabled = False
        
    End If
    
End Sub

Private Sub cmd_slab_init_Click()

    Dim iCnt As Long
    Dim iRow As Integer
    
    For iCnt = 1 To iSlab_cnt
        lbl_slab(iCnt).Caption = "删除"
    Next iCnt
    
    Call cmd_slab_del_Click
    
    iSlab_cnt = 0
    lbl_max.Caption = ""
    lbl_min.Caption = ""
    sdb_charge_wgt.Value = 0
    cmd_slab_init.Enabled = False
    cmd_slab_del.Enabled = False
    cmd_slab_complete.Enabled = False
    
    Call Form_Ref
    
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
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
    SS1.RetainSelBlock = False
    SS1.OperationMode = OperationModeNormal
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    txt_line.Text = "1"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim iCount As Integer
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    Dim iCnt As Long
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        
        SS1.SetFocus
        
        txt_plt.Text = "B1"
        Call txt_plt_KeyUp(0, 0)
        txt_line.Text = "1"
        
        For iCnt = 1 To iSlab_cnt
            Unload lbl_slab(iCnt)
        Next iCnt
    
        iSlab_cnt = 0
        lMain_row = 0
        cmd_slab_del.Enabled = False
        cmd_slab_complete.Enabled = False
        
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sQuery As String
    Dim dValue As String
    
    Dim iCnt As Long
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc").Item("Spread"))
        SS1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        iSlab_cnt = 0
        lMain_row = 0
        sdb_charge_wgt.Value = 0
        sdb_heat_edt_seq.Value = 0
        txt_MLT_PROC_CD.Text = ""
        lbl_max.Caption = ""
        lbl_min.Caption = ""
    End If
            
    Exit Sub

Refer_Err:

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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub lbl_slab_DblClick(Index As Integer)

    Dim sSeq As String
    
    If Index < 10 Then
        sSeq = "0" & Index
    Else
        sSeq = Trim(str(Index))
    End If
    
    If lbl_slab(Index).BackColor = &HC0C0FF Then
        If lbl_slab(Index).Tag = "H" Then
            lbl_slab(Index).BackColor = &H8080FF
            lbl_slab(Index).ForeColor = &HFF0000
        Else
            lbl_slab(Index).BackColor = &HFF8080
            lbl_slab(Index).ForeColor = &HFF0000
        End If
        
        lbl_slab(Index).Caption = sSeq
    Else
        lbl_slab(Index).BackColor = &HC0C0FF
        lbl_slab(Index).ForeColor = &HFF0000
        lbl_slab(Index).Caption = "删除"
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim sTemp_ord As String
    Dim sSeq As String
    Dim iRow As Integer
    Dim iCnt As Long
    Dim dWgt As Double
    Dim dLen As Double
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    'UPDATE AUTHORITY
    If Mid(sAuthority, 3, 1) <> "1" Then
        Call Gp_MsgBoxDisplay("It is no HMI 炉次编制 authority", "I")
    End If
    
    If SS1.MaxRows < 1 Or Row < 1 Then Exit Sub
    
    SS1.Row = Row
    SS1.Col = 0
    
    If SS1.Text <> "选择" Then
    
        If iSlab_cnt = 0 Then
            'txt_mlt_proc_cd
            SS1.Col = 9
            txt_MLT_PROC_CD.Text = SS1.Text
            
            lMain_row = Row
            
            Call Max_wgt(txt_MLT_PROC_CD.Text)
            Call Min_wgt(txt_MLT_PROC_CD.Text)
            Call Max_heat_edt_seq
            
        Else
            If Condition_Compare(Row) = False Then Exit Sub
            
        End If
        
        SS1.Col = 8
        If sdb_charge_wgt.Value + SS1.Value > Val(lbl_max.Caption) Then
            Call Gp_MsgBoxDisplay("炉次重量 > Max 重量")
            Exit Sub
        End If
        
        SS1.Col = 0
        SS1.Text = "选择"
        
        iSlab_cnt = iSlab_cnt + 1
        
        If iSlab_cnt < 10 Then
            sSeq = "0" & iSlab_cnt
        Else
            sSeq = Trim(str(iSlab_cnt))
        End If
        
        'SLAB_LEN
        SS1.Col = 7
        dLen = SS1.Value
        
        'SLAB_WGT
        SS1.Col = 8
        sdb_charge_wgt.Value = sdb_charge_wgt.Value + SS1.Value
        dWgt = SS1.Value
        
        Load lbl_slab(iSlab_cnt)
        
        lbl_slab(iSlab_cnt).Caption = sSeq
        lbl_slab(iSlab_cnt).Height = (3090 / Val(lbl_max.Caption)) * SS1.Value
        lbl_slab(iSlab_cnt).Width = (4605 / 20000) * dLen
            
        If iSlab_cnt = 1 Then
            lbl_slab(iSlab_cnt).Top = 8595 - lbl_slab(iSlab_cnt).Height
        Else
            lbl_slab(iSlab_cnt).Top = lbl_slab(iSlab_cnt - 1).Top - lbl_slab(iSlab_cnt).Height
        End If
        
        'HCR
        SS1.Col = 15
        lbl_slab(iSlab_cnt).Tag = SS1.Text
        
        If SS1.Text = "H" Then
            lbl_slab(iSlab_cnt).BackColor = &H8080FF
        Else
            lbl_slab(iSlab_cnt).BackColor = &HFF8080
        End If
        
        lbl_slab(iSlab_cnt).Visible = True
        
        'HEAT_EDT_SEQ
        SS1.Col = 2
        SS1.Value = sdb_heat_edt_seq.Value
        
        'HEAT_SLQB_SEQ
        SS1.Col = 3
        SS1.Text = sSeq
        
        Call Gp_Sp_BlockColor(SS1, 1, SS1.MaxCols, Row, Row, , &HFFFF80)
        
        'EP_SLAB_EDT  TABLE HEAT_EDT_SEQ, HEAT_SLSB_SEQ UPDATEING
        SS1.Col = 1 'SLAB_EDT_SEQ
        Call Slab_Seq_Create("U", SS1.Value, sSeq)
        
        cmd_slab_init.Enabled = True
        cmd_slab_del.Enabled = True
        cmd_slab_complete.Enabled = True
    
    Else
    
        'ss1.Text = ""
        
        'SLAB_WGT
        'ss1.Col = 8
        'sdb_charge_wgt.Value = sdb_charge_wgt.Value - ss1.Value
        
        'Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
    
    End If
        
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.SS1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub txt_plt_DblClick()

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

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_prod_cd.Text)) = txt_prod_cd.MaxLength Then
        txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
    Else
        txt_prod_cd_name.Text = ""
    End If
    
End Sub

Public Sub Max_wgt(sProc_cd As String)

    Dim sQuery As String
    
    'Max Wgt
    sQuery = "SELECT  NVL(MAXI,0) "
    sQuery = sQuery + "  From NISCO.EP_CHARGE_S2 "
    sQuery = sQuery + " Where PLT              = '" + txt_plt.Text + "' "
    sQuery = sQuery + "   AND PRC_LINE         = '" + txt_line.Text + "' "
    sQuery = sQuery + "   AND SUBSTR(MLT_TOT_PROC_CD,3," & Len(txt_MLT_PROC_CD.Text) & ")  = '" + txt_MLT_PROC_CD.Text + "' "
    
    lbl_max.Caption = str(Gf_FloatFind(M_CN1, sQuery))
        
End Sub

Public Sub Min_wgt(sOrderNo As String)

    Dim sQuery As String
    
    'Min Wgt
    sQuery = "SELECT  NVL(MINI,0) "
    sQuery = sQuery + "  From NISCO.EP_CHARGE_S2 "
    sQuery = sQuery + " Where PLT              = '" + txt_plt.Text + "' "
    sQuery = sQuery + "   AND PRC_LINE         = '" + txt_line.Text + "' "
    sQuery = sQuery + "   AND SUBSTR(MLT_TOT_PROC_CD,3," & Len(txt_MLT_PROC_CD.Text) & ")  = '" + txt_MLT_PROC_CD.Text + "' "
    
    lbl_min.Caption = str(Gf_FloatFind(M_CN1, sQuery))
        
End Sub

Public Sub Max_heat_edt_seq()

    Dim sQuery As String
    
    'Min Wgt
    sQuery = "SELECT  NVL(MAX(HEAT_EDT_SEQ),0) "
    sQuery = sQuery + "  From NISCO.EP_SLAB_EDT "
    
    sdb_heat_edt_seq.Value = Gf_FloatFind(M_CN1, sQuery) + 1
        
End Sub


Private Sub Slab_Seq_Create(iType As String, SLAB_EDT_SEQ As Variant, Seq As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    'HEAT_EDT_SEQ, HEAT_SLAB_SEQ
    sQuery = "{call AEB1080C.P_MODIFY ('" + iType + "'," & SLAB_EDT_SEQ & "," & sdb_heat_edt_seq.Value & ",'" + Seq + "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)

End Sub

Private Function Condition_Compare(iRow As Long) As Boolean

    Dim sTemp As String
    Dim dTemp As Double
    
    Condition_Compare = True
    
    'STLGRD
    SS1.Row = lMain_row
    SS1.Col = 4
    sTemp = SS1.Text
    SS1.Row = iRow
    
    If sTemp <> SS1.Text Then
        Call Gp_MsgBoxDisplay("钢种不一致")
        Condition_Compare = False
        Exit Function
    End If
    
    'THK
    SS1.Row = lMain_row
    SS1.Col = 5
    dTemp = SS1.Value
    SS1.Row = iRow
    
    If dTemp <> SS1.Value Then
        Call Gp_MsgBoxDisplay("厚度不一致")
        Condition_Compare = False
        Exit Function
    End If
    
    'WID
    SS1.Row = lMain_row
    SS1.Col = 6
    dTemp = SS1.Value
    SS1.Row = iRow
    
    If dTemp <> SS1.Value Then
        Call Gp_MsgBoxDisplay("宽度不一致")
        Condition_Compare = False
        Exit Function
    End If
    
    'MLT_PROC_CD
    SS1.Row = lMain_row
    SS1.Col = 9
    sTemp = SS1.Text
    SS1.Row = iRow
    
    If sTemp <> SS1.Text Then
        Call Gp_MsgBoxDisplay("工序流程不一致")
        Condition_Compare = False
        Exit Function
    End If
    
    'PROD_CD
    SS1.Row = lMain_row
    SS1.Col = 10
    sTemp = SS1.Text
    SS1.Row = iRow
    
    If sTemp <> SS1.Text Then
        Call Gp_MsgBoxDisplay("产品不一致")
        Condition_Compare = False
        Exit Function
    End If
    
End Function
