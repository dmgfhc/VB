VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Begin VB.Form AAA1030C 
   Caption         =   "销售计划录入_AAA1030C"
   ClientHeight    =   9885
   ClientLeft      =   330
   ClientTop       =   3030
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txt_prod_cd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AAA1030C.frx":0000
      Left            =   6900
      List            =   "AAA1030C.frx":000D
      TabIndex        =   16
      Tag             =   "产品代码"
      Top             =   135
      Width           =   780
   End
   Begin CSTextLibCtl.sidbEdit sdb_tot 
      Height          =   330
      Left            =   1485
      TabIndex        =   11
      Top             =   1305
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Modified        =   0   'False
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
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand SCmd2 
      Height          =   465
      Left            =   9570
      TabIndex        =   8
      Top             =   105
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "炼钢工序计算"
   End
   Begin Threed.SSCommand SCmd1 
      Height          =   465
      Left            =   9570
      TabIndex        =   7
      Top             =   705
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "计划汇总"
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   930
      Left            =   6525
      TabIndex        =   6
      Top             =   1980
      Width           =   8730
      _Version        =   393216
      _ExtentX        =   15399
      _ExtentY        =   1640
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
      MaxCols         =   6
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SpreadDesigner  =   "AAA1030C.frx":001D
   End
   Begin VB.TextBox txt_cust_name 
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
      Left            =   2955
      MaxLength       =   40
      TabIndex        =   2
      Top             =   495
      Width           =   4725
   End
   Begin VB.TextBox txt_cust_cd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1485
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "客户"
      Top             =   495
      Width           =   1470
   End
   Begin VB.TextBox txt_stlgrd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1485
      MaxLength       =   11
      TabIndex        =   3
      Tag             =   "钢种"
      Top             =   855
      Width           =   1470
   End
   Begin VB.TextBox txt_stlgrd_des 
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
      Left            =   2955
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   855
      Width           =   4725
   End
   Begin InDate.ULabel ULabel6 
      Height          =   285
      Left            =   135
      Top             =   855
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel5 
      Height          =   300
      Left            =   5565
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
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
   Begin InDate.UDate dtp_date_str 
      Height          =   300
      Left            =   1485
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   135
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Left            =   135
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "日期"
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
      ForeColor       =   0
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   6015
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   10610
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
      MaxCols         =   0
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1030C.frx":058F
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   135
      Top             =   495
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "客户"
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
   Begin FPSpread.vaSpread ss3 
      Height          =   1200
      Left            =   120
      TabIndex        =   9
      Top             =   1710
      Width           =   6315
      _Version        =   393216
      _ExtentX        =   11139
      _ExtentY        =   2117
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
      MaxCols         =   4
      MaxRows         =   3
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SpreadDesigner  =   "AAA1030C.frx":07A8
   End
   Begin FPSpread.vaSpread ss4 
      Height          =   660
      Left            =   6525
      TabIndex        =   10
      Top             =   1305
      Width           =   8730
      _Version        =   393216
      _ExtentX        =   15399
      _ExtentY        =   1164
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
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
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SpreadDesigner  =   "AAA1030C.frx":0C34
   End
   Begin InDate.ULabel ULabel4 
      Height          =   330
      Left            =   3810
      Top             =   1305
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "已使用(h)"
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
      Height          =   330
      Left            =   135
      Top             =   1305
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "总时间(h)"
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
   Begin CSTextLibCtl.sidbEdit sdb_work 
      Height          =   330
      Left            =   5160
      TabIndex        =   12
      Top             =   1305
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Modified        =   0   'False
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.UDate dtp_copy_to 
      Height          =   330
      Left            =   12645
      TabIndex        =   13
      Top             =   690
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel9 
      Height          =   330
      Left            =   11520
      Top             =   690
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Caption         =   "复制到"
      Alignment       =   1
      BackColor       =   16777088
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
   Begin InDate.UDate dtp_copy_from 
      Height          =   330
      Left            =   12645
      TabIndex        =   14
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel7 
      Height          =   330
      Left            =   11520
      Top             =   240
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Caption         =   "从"
      Alignment       =   1
      BackColor       =   16777088
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   330
      Left            =   13980
      TabIndex        =   15
      Top             =   480
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "复制"
   End
   Begin VB.Shape Shape1 
      Height          =   1050
      Left            =   11400
      Top             =   120
      Width           =   3765
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   90
      X2              =   15210
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   90
      X2              =   15255
      Y1              =   1260
      Y2              =   1260
   End
End
Attribute VB_Name = "AAA1030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       production plan
'-- Sub_System Name
'-- Program Name
'-- Program ID        AAA1030C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2003.7.9
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim THK_GRP As Collection
Dim WID_GRP As Collection
Dim MIN_VALUE As Collection
Dim MAX_VALUE As Collection

Dim arrValue As Variant

Private Sub Form_Define()
    
    Dim sQuery As String
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(dtp_date_str, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_cust_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_stlgrd_des, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Spread_Collection
    Sc1.Add Item:="AAA1030C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub SCmd2_Click()

On Error GoTo SCmd2_Error

    Dim sQuery As String
    Dim iCount As Integer
    
    'If dtp_date_str.Enabled Then Exit Sub
    
    Dim adoCmd As ADODB.Command
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    For iCount = 1 To 7
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount), adVariant, adParamOutput)
    Next iCount
    
    'CAST
    sQuery = "{call AAA3011P ('" + Left(dtp_date_str.RawData, 6) + "', '" + txt_prod_cd.Text + "', 'B1', 'BF', ?,?,?,?,?,?,? )}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(6) <> "" Then
        Call Gp_MsgBoxDisplay(adoCmd(6))
        Set adoCmd = Nothing
        Exit Sub
    End If
    
    With ss2
        .Row = 1
        .Col = 2
        .Value = adoCmd(0)
        .Col = 3
        .Value = adoCmd(4)
        .Col = 4
        .Value = adoCmd(1)
        .Col = 5
        .Value = adoCmd(3)
        .Col = 6
        .Value = adoCmd(2)
    End With
    
    'BOF
    sQuery = "{call AAA3012P ('" + Left(dtp_date_str.RawData, 6) + "', '" + txt_prod_cd.Text + "', 'B1', 'BC', ?,?,?,?,?,?,? )}"
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(6) <> "" Then
        Call Gp_MsgBoxDisplay(adoCmd(6))
        Set adoCmd = Nothing
        Exit Sub
    End If
    
    With ss2
        .Row = 2
        .Col = 2
        .Value = adoCmd(0)
        .Col = 3
        .Value = adoCmd(4)
        .Col = 4
        .Value = adoCmd(1)
        .Col = 5
        .Value = adoCmd(3)
        .Col = 6
        .Value = adoCmd(2)
    End With
        
    Set adoCmd = Nothing
    
    If subCollectionAdd = True Then
        Call subValueCheck
    End If
    
    Exit Sub

SCmd2_Error:

    Call Gp_MsgBoxDisplay("炼钢能力检查 Error : " & Error)

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Menu_Setting

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
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Sp_Setting2(ss2)
    Call Sp_Setting2(ss3)
    Call Sp_Setting2(ss4)
    
    Call Sp_Setting1
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    txt_prod_cd.Text = "PP"

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set THK_GRP = Nothing
    Set WID_GRP = Nothing
    Set MIN_VALUE = Nothing
    Set MAX_VALUE = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    SCmd2.Enabled = False
    
    rControl(1).SetFocus
    
    sdb_tot.Value = 0
    sdb_work.Value = 0
    ss1.MaxCols = 0
    ss1.MaxRows = 0
    
    ss2.ClearRange 1, 1, ss2.MaxCols, ss2.MaxRows, False
    ss3.ClearRange 1, 1, ss3.MaxCols, ss3.MaxRows, False
    ss4.ClearRange 1, 1, ss4.MaxCols, ss4.MaxRows, False
 
End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
        
        If Sp_Header_Refer() Then
            If Sp_Data_Refer() Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Menu_Setting
                Call Gp_Ms_ControlLock(Mc1!lControl, True)
                If Left(dtp_date_str.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
                    Call Gp_Sp_BlockLock(ss1, 1, -1, 1, -1, True)
                    SCmd2.Enabled = False
                Else
                    SCmd2.Enabled = True
                End If
            Else
                If Left(dtp_date_str.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
                    Call Gp_Sp_BlockLock(ss1, 1, -1, 1, -1, True)
                    SCmd2.Enabled = False
                Else
                    SCmd2.Enabled = True
                End If
            End If
            Call Sp_Other_Refer
        End If
    Else
        sMesg = sMesg + " 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
    txt_cust_cd.Enabled = True
    txt_stlgrd.Enabled = True
    
End Sub

Public Sub Form_Pro()

    If Left(dtp_date_str.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
        Call Gp_MsgBoxDisplay("Can't Process...")
        Exit Sub
    End If
    
'    If subCollectionAdd = True Then
'        Call subValueCheck
'    End If
    
    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
        Call ss2_Process
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Form_Ref
    End If
    
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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

Public Sub Spread_Del()
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Visible = False
        MDIMain.Line1.Visible = False
        
        PopupMenu MDIMain.PopUp_Spread
        
        MDIMain.Mnu_Sorting.Visible = True
        MDIMain.Line1.Visible = True
    End If

End Sub

Public Sub Sp_Setting1()

    With ss1

        .ColHeaderRows = 3
        .RowHeaderCols = 2
        .Col = -1
        .Row = SpreadHeader + 1
        .FontBold = True
        
        .RowHeight(SpreadHeader) = 12
        .RowHeight(SpreadHeader + 1) = 12
        
        .Row = SpreadHeader + 2
        
        .ColWidth(0) = 10
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
          
        .Row = SpreadHeader
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        .Row = SpreadHeader + 1
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        
        .Row = SpreadHeader + 2
        .RowHidden = True
        
        .Col = SpreadHeader + 1
        .ColHidden = True
        
    End With

End Sub

Public Sub Menu_Setting()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel
    
End Sub

Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    Dim sQuery2 As String
    
    Dim AdoRs2 As ADODB.Recordset
    Dim ArrayRecords2 As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "SELECT THK_CD, FR_THK, TO_THK "
    sQuery = sQuery + "   FROM BP_THICK_GRP "
    sQuery = sQuery + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery = sQuery + "    AND THK_CD <> '*' "
    sQuery = sQuery + "  ORDER BY THK_CD "
    
    With ss1

        Sp_Header_Refer = True
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
        Set adoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 2
        
            For iCol = 0 To .MaxCols - 1 Step 2
            
               .Col = iCol + 1
               .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                  
                .Col = iCol + 2
                .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                           
                .Col = iCol + 1:  .Row = SpreadHeader + 1:  .Text = "实绩"
                .Col = iCol + 2:  .Row = SpreadHeader + 1:  .Text = "计划"
                
                .Col = iCol + 1
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                .Col = iCol + 2
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                'Column Type Setting
                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                .ColWidth(iCol + 1) = 9
                
                .Col = iCol + 2: .Col2 = iCol + 2
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                .ColWidth(iCol + 2) = 9
                
                iCnt = iCnt + 1
                
            Next iCol
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Set AdoRs2 = New ADODB.Recordset
    
    sQuery2 = "SELECT WID_CD, FR_WID, TO_WID "
    sQuery2 = sQuery2 + "   FROM BP_WIDTH_GRP "
    sQuery2 = sQuery2 + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery2 = sQuery2 + "    AND WID_CD <> '*' "
    sQuery2 = sQuery2 + "  ORDER BY WID_CD "
    
    With ss1

        Sp_Header_Refer = True
        .ColWidth(0) = 15
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs2.Open sQuery2, M_CN1, adOpenKeyset
        
        If AdoRs2.BOF Or AdoRs2.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            AdoRs2.Close
            Set AdoRs2 = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords2 = AdoRs2.GetRows
        AdoRs2.Close
        Set AdoRs2 = Nothing

        If UBound(ArrayRecords2, 2) + 1 <> 0 Then
        
            .MaxRows = (UBound(ArrayRecords2, 2) + 1)
            iCnt = 0
            
            For iRow = 1 To .MaxRows
            
                .Row = iRow
                .Col = SpreadHeader
                
                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords2(1, iCnt)) & " ~ " & Trim(ArrayRecords2(2, iCnt)) & "mm"
                End If
                
                .Col = SpreadHeader + 1
                .Text = Trim(ArrayRecords2(0, iCnt))
                
                .Row = iRow + 2: .Row2 = iRow + 2
                .Col = 1: .Col2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                iCnt = iCnt + 1
            Next iRow
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    With ss1
    
        For iCol = 1 To .MaxCols - 1 Step 2
            .Col = iCol
            .Row = 1
            .Col2 = iCol
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            .Protect = True
        Next iCol
        
        For iCol = 2 To .MaxCols Step 2
            .Col = iCol
            .Row = 1
            .Col2 = iCol
            .Row2 = .MaxRows
             If Trim(txt_prod_cd.Text) = "" Or Trim(txt_cust_cd.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then
                .BlockMode = True
                .Lock = True
                .BlockMode = False
                .Protect = True
             Else
                .BlockMode = True
                .Lock = False
                .BackColor = &HC0FFFF
                .BlockMode = False
                .Protect = True
             End If
        Next iCol
        
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Set AdoRs2 = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim sWID_GRP As String
    Dim sTHK_GRP As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    sEdate = Left(dtp_date_str.RawData, 6)
  
    sQuery = "SELECT WID_GRP, THK_GRP, sum(RST_WGT),sum(PLN_WGT)"
    sQuery = sQuery + "   FROM AP_SALES_PLAN "
    sQuery = sQuery + "  WHERE YEAR_MONTH =      '" + sEdate + "' "
    sQuery = sQuery + "    AND CUST_CD    LIKE   '" + Trim(txt_cust_cd.Text) + "%' "
    sQuery = sQuery + "    AND PROD_CD    LIKE   '" + Trim(txt_prod_cd.Text) + "%' "
    sQuery = sQuery + "    AND STLGRD     LIKE   '" + Trim(txt_stlgrd.Text) + "%' "
    sQuery = sQuery + "  GROUP BY WID_GRP, THK_GRP "
    sQuery = sQuery + "  ORDER BY WID_GRP, THK_GRP "
    
    With ss1

        Sp_Data_Refer = True
        
        .ReDraw = False
       ' .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
     '   Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
            iRow = 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .Row = iRow
                .Col = SpreadHeader + 1
                 sWID_GRP = .Text
                 Do While iRow <= .MaxRows And sWID_GRP <> Trim(ArrayRecords(0, iCnt))
                    iRow = iRow + 1
                    .Row = iRow
                    sWID_GRP = .Text
                 Loop
                           
                 For iCol = 1 To .MaxCols - 1 Step 2
                    .Col = iCol
                    .Row = SpreadHeader + 2
                    sTHK_GRP = .Text

                    If sTHK_GRP = ArrayRecords(1, iCnt) Then
                        
                        .Row = iRow
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Or ArrayRecords(2, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(2, iCnt))
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Or ArrayRecords(3, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(3, iCnt))
                        End If
                
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
        Screen.MousePointer = vbDefault
        
    End With
         
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
    
    If Trim(txt_prod_cd.Text) = "" Or Trim(txt_cust_cd.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then
       Sp_Process = False
       Call Gp_MsgBoxDisplay("不能保存 ...")
       Exit Function
    End If
    
    With ss1
    
        'MaxRow = 0 is Exit Function Or iCount = 0
        If .MaxRows < 1 Then
            Sp_Process = False
            Exit Function
        End If
        
        Screen.MousePointer = vbHourglass
        
        .ReDraw = False
        
        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Sp_Process = False: Exit Function
        End If
        
        'Ado Setting
        Conn.CursorLocation = adUseServer
        Set adoCmd = New ADODB.Command
        
        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = Sc.Item("P-M")
        
        Conn.BeginTrans
        
        'Ceate Parameter (Input) iType + iColumn
        For iCount = 1 To 8
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            'Parameters Setting
            For iCol = 2 To .MaxCols Step 2
            
                .Col = iCol
                If Trim(.Text) <> "" Then
                
                    .Row = SpreadHeader + 2
                    .Col = iCol
                    adoCmd.Parameters(4).Value = .Text     'thk_grp
              
                    .Row = iRow
                    .Col = SpreadHeader + 1
                    adoCmd.Parameters(5).Value = .Text     'wid_grp
                    
                    .Col = iCol
                 
                    If Trim(.Text) = "" Then               'plan_value
                        adoCmd.Parameters(6).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(6).Value = dTempInt
                    End If
                    
                    adoCmd.Parameters(7).Value = sUserID                            'User-id
                    
                    adoCmd.Parameters(0).Value = Mid(dtp_date_str.Text, 1, 4) + _
                                                 Mid(dtp_date_str.Text, 6, 2)        'YEAR_MONTH
                    adoCmd.Parameters(1).Value = txt_cust_cd.Text                    'CUST_CD
                    adoCmd.Parameters(2).Value = txt_prod_cd.Text                    'PROD_CD
                    adoCmd.Parameters(3).Value = txt_stlgrd.Text                     'STLGRD
                                   
                    adoCmd.Execute
                    
                    'Error Check
                    If adoCmd("Error") <> "0" Then
               
                        ret_Result_ErrCode = adoCmd("Error")
                        ret_Result_ErrMsg = adoCmd("Messg")
                        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
               
                        Call Gp_MsgBoxDisplay(sErrMessg)
                        Screen.MousePointer = vbDefault
                        Set adoCmd = Nothing
                        Conn.RollbackTrans
                        Sp_Process = False
                        Exit Function
               
                     End If
                
                End If
            
            Next iCol
            
        Next iRow
        
        Conn.CommitTrans
        .ReDraw = True
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据修改完成"
        Screen.MousePointer = vbDefault
        Exit Function
    
    End With

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("SpreadPro_Error : " & Error)

End Function

Private Sub SCmd1_Click()

    Load AAA1080C
    AAA1080C.txt_year_month.Text = Left(dtp_date_str.RawData, 6)
    AAA1080C.Show 1
    
End Sub

Private Sub SSCommand1_Click()

'   Dim sQuery As String
'   If dtp_copy_from.RawData <> "" And dtp_copy_to.RawData <> "" And dtp_copy_to.RawData > dtp_copy_from.RawData Then
'      sQuery = "{call AAA1030C.P_MODIFY1('" + dtp_copy_from.RawData + "','" + dtp_copy_to.RawData + "','" + sUserID + "',?,?)}"
'
'   End If
On Error GoTo Cmd1_Error

    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim sMesg As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command

    If dtp_copy_from.RawData = "" Or dtp_copy_to.RawData = "" Or dtp_copy_to.RawData <= dtp_copy_from.RawData Then
        Call Gp_MsgBoxDisplay("必须输入正确的日期...")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
        
    'Db Connection Check
'    If GF_DbConnect = False Then
'       Exit Sub
'    End If
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = M_CN1
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "AAA1030C.P_MODIFY1"
    
    M_CN1.BeginTrans
    
    'Ceate Parameter (Input) iType + iColumn
    For iCount = 1 To 3
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
    adoCmd.Parameters(0).Value = dtp_copy_from.RawData
    adoCmd.Parameters(1).Value = dtp_copy_to.RawData
    adoCmd.Parameters(2).Value = sUserID                            'User-id
    adoCmd.Execute
     
     'Error Check
     If adoCmd("Error") <> "0" Then

         ret_Result_ErrCode = adoCmd("Error")
         ret_Result_ErrMsg = adoCmd("Messg")
         sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg

         Call Gp_MsgBoxDisplay(sErrMessg)
         Screen.MousePointer = vbDefault
         Set adoCmd = Nothing
         M_CN1.RollbackTrans
         Exit Sub

      End If
        
      M_CN1.CommitTrans
      Screen.MousePointer = vbDefault
      Exit Sub
    
Cmd1_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    M_CN1.RollbackTrans
    Call Gp_MsgBoxDisplay("Cmd1_Error : " & Error)

End Sub

Private Sub txt_cust_cd_DblClick()
    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_name

        DD.nameType = "1"
        Call Gf_Customer_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_name.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     
End Sub

Private Sub txt_stlgrd_DblClick()
    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_des

        DD.nameType = "2"
        Call Gf_Stlgrd_DD_AC(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
        txt_stlgrd_des.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
    Else
        txt_stlgrd_des.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Function subCollectionAdd() As Boolean

    Dim i As Integer
    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    If dtp_date_str.RawData = "" Or Trim(txt_prod_cd.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then Exit Function

    Set adoRs = New ADODB.Recordset

'Collection Clear ----------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
    sQuery = "SELECT THK_GRP , WID_GRP , SUM(NVL(MIN,0)) , SUM(NVL(MAX,0)) FROM AP_LIMIT_CON WHERE "
    sQuery = sQuery + "YEAR_MONTH = '" + Left(dtp_date_str.RawData, 6) + "' AND "
    sQuery = sQuery + "PROD_CD    = '" + txt_prod_cd.Text + "' AND "
    sQuery = sQuery + "STLGRD     = '" + txt_stlgrd.Text + "' "
    sQuery = sQuery + "GROUP BY THK_GRP, WID_GRP "
    
    adoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If adoRs.BOF Or adoRs.EOF Then
        adoRs.Close
        Set adoRs = Nothing
        Exit Function
    End If
    
    arrValue = adoRs.GetRows
    adoRs.Close
    Set adoRs = Nothing
    subCollectionAdd = True
                
End Function

Private Sub subValueCheck()

    Dim i As Integer
    Dim j As Integer
    Dim bCheck As Boolean
    Dim iCnt As Integer
    Dim sMesg As String
    
    With ss1
        
        For i = 1 To .MaxRows
        
            For j = 2 To .MaxCols Step 2
                If FunValueCheck(Gf_Get_Cell_Text(SpreadHeader + 2, j), Gf_Get_Cell_Text(i, SpreadHeader + 1), GF_GET_CELL_VALUE(i, j)) = False Then
                    Call subCellBackColor(j, i)
                    iCnt = iCnt + 1
                End If
            Next j
        
        Next i
    
        If iCnt > 0 Then
            sMesg = "超出生产限制条件"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    End With

End Sub

Private Function FunValueCheck(ByVal sTHK_GRP As String, ByVal sWID_GRP As String, ByVal dValue As Double) As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    If UBound(arrValue, 2) < 0 Then Exit Function
    
    For i = 0 To UBound(arrValue, 2)
    
        'If dValue = 310 Then
         '   Exit Function
        'End If
        
        If arrValue(0, i) = sTHK_GRP And arrValue(1, i) = sWID_GRP Then
        
            If arrValue(2, i) >= 0 And arrValue(3, i) <> 0 Then
                
                If arrValue(2, i) <= dValue And dValue <= arrValue(3, i) Then
                    FunValueCheck = True
                Else
                    FunValueCheck = False
                End If
                Exit Function
                            
            End If
        
        End If
    Next i
    
    FunValueCheck = True
    
End Function

Private Sub subCellBackColor(ByVal iCol As Integer, ByVal iRow As Integer)
    
    With ss1
        .Col = iCol
        .Row = iRow
        .BackColor = vbRed
    End With
    
End Sub

Private Function GF_GET_CELL_VALUE(ByVal iRow As Long, ByVal iCol As Long) As Variant
    
    With ss1
        .Row = iRow
        .Col = iCol
        GF_GET_CELL_VALUE = IIf(.Value = "", 0, .Value)
    End With
    
End Function

Private Function Gf_Get_Cell_Text(ByVal iRow As Long, ByVal iCol As Long) As Variant
    
    With ss1
        .Row = iRow
        .Col = iCol
        Gf_Get_Cell_Text = .Text
    End With
    
End Function

Public Sub Sp_Setting2(ByVal sPname As Variant)

    With sPname
    
        .RowHeight(-1) = 12
        .RowHeight(0) = 16
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .OperationMode = OperationModeRead
        '.RetainSelBlock = True

        .UserResize = UserResizeNone
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
        
        .Col = 0
        .Row = -1
        .FontBold = True
        
        .LockBackColor = RGB(255, 255, 255)
        
        If .Name = "ss3" Then Call Gp_Sp_RowColor(ss3, 3, vbRed)
        If .Name = "ss4" Then .RowHeadersShow = False
        
    End With
    
End Sub

Public Function Sp_Other_Refer() As Boolean

On Error GoTo Sp_Other_Refer_Error

    Dim iCol As Integer
    Dim lSum As Long
    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "{ call AAA1030C.P_REFER('" + Left(dtp_date_str.RawData, 6) + "') }"
    
    adoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If adoRs.BOF Or adoRs.EOF Then
    
        Sp_Other_Refer = False
        adoRs.Close
        Set adoRs = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If
        
    ArrayRecords = adoRs.GetRows
    
    adoRs.Close
    Set adoRs = Nothing
    
    sdb_tot.Value = ArrayRecords(0, 0)

    ss3.Row = 1
    For iCol = 1 To 4
        ss3.Col = iCol
        ss3.Value = Trim(ArrayRecords(iCol, 0))
    Next iCol
    
    ss3.Row = 2
    For iCol = 1 To 4
        ss3.Col = iCol
        ss3.Value = Trim(ArrayRecords(iCol + 4, 0))
    Next iCol

    ss3.Row = 3
    For iCol = 2 To 4
        ss3.Col = iCol
        ss3.Value = Trim(ArrayRecords(iCol, 0) + ArrayRecords(iCol + 4, 0))
        If iCol = 4 Then sdb_work.Value = Trim(ArrayRecords(iCol, 0) + ArrayRecords(iCol + 4, 0))
    Next iCol

    lSum = 0
    ss4.Row = 1
    For iCol = 1 To 4
        ss4.Col = iCol
        ss4.Value = Trim(ArrayRecords(iCol + 8, 0))
'        lSum = lSum + Trim(ArrayRecords(iCol + 8, 0))
    Next iCol
'    ss4.Col = 4
'    ss4.Value = lSum
    Call Gp_Sp_CellColor(ss4, 4, 1, vbRed)

    ss2.Row = 1
    For iCol = 1 To 6
        ss2.Col = iCol
        ss2.Value = Trim(ArrayRecords(iCol + 12, 0))
    Next iCol

    ss2.Row = 2
    For iCol = 1 To 6
        ss2.Col = iCol
        ss2.Value = Trim(ArrayRecords(iCol + 18, 0))
    Next iCol
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Screen.MousePointer = vbDefault
    Exit Function

Sp_Other_Refer_Error:
    
    Set adoRs = Nothing
    Sp_Other_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Sp_Other_Refer_Error : " & Error)
    
End Function

Public Sub ss2_Process()

On Error GoTo ss2_Process_Error

    Dim sQuery As String
    Dim iCount As Integer
    
    If txt_prod_cd.Text <> "HC" And txt_prod_cd.Text <> "PP" Then Exit Sub
    
    Dim adoCmd As ADODB.Command
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    For iCount = 1 To 7
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount), adVariant, adParamOutput)
    Next iCount
    
    'CAST
    sQuery = "{call AAA3010P ('" + Left(dtp_date_str.RawData, 6) + "', '" + txt_prod_cd.Text + "', 'C1', 'CB', ?,?,?,?,?,?,? )}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(6) <> "" Then
        Call Gp_MsgBoxDisplay(adoCmd(6))
        Set adoCmd = Nothing
        Exit Sub
    End If
    
    Set adoCmd = Nothing
    Exit Sub

ss2_Process_Error:

    Set adoCmd = Nothing
    Call Gp_MsgBoxDisplay(Error)
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Stlgrd_DD_AC
'   2.Name         : Stlgrd Code Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .20
'   7.Modify Date  :
'   8.Comment      : Stlgrd Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Stlgrd_DD_AC(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "S"        'Stlgrd Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.rControl.Item(1).Text) & "%' AND STLGRD_FL <> 'H'  "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.sPname.Text) & "%' AND STLGRD_FL <> 'H' "
            
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
   
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                sNew_Name = DD.sPname.Text
            End If
            
            DD.sPname.TabStop = True
            DD.sPname.SetFocus
            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
            DD.sPname.EditMode = True
            DD.sPname.TabStop = False
            
            If DD.sSelect Then
                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
            End If
            
        End If
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function
