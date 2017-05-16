VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQT0040C 
   Caption         =   "Test CertPrint"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Begin VB.TextBox txt_PROD_CD 
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
      Left            =   5745
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "产品代码"
      Top             =   90
      Width           =   540
   End
   Begin VB.TextBox txt_CUST_NAME 
      Enabled         =   0   'False
      Height          =   310
      Left            =   3000
      TabIndex        =   6
      Top             =   555
      Width           =   3285
   End
   Begin VB.TextBox txt_ORD_NO 
      Height          =   310
      Left            =   7650
      MaxLength       =   11
      TabIndex        =   5
      Top             =   555
      Width           =   1305
   End
   Begin VB.TextBox txt_CUST_CD 
      Height          =   310
      Left            =   1860
      MaxLength       =   6
      TabIndex        =   4
      Top             =   555
      Width           =   1125
   End
   Begin VB.TextBox txt_CERT_NO 
      Height          =   310
      Left            =   1860
      MaxLength       =   14
      TabIndex        =   3
      Top             =   90
      Width           =   2235
   End
   Begin VB.TextBox txt_SHP_ISP_NO 
      Height          =   310
      Left            =   12270
      MaxLength       =   11
      TabIndex        =   2
      Top             =   90
      Width           =   1305
   End
   Begin VB.TextBox TXT_PONO 
      Height          =   310
      Left            =   12270
      TabIndex        =   1
      Top             =   570
      Width           =   1410
   End
   Begin VB.TextBox txt_SAVE_DIR 
      Enabled         =   0   'False
      Height          =   310
      Left            =   6540
      TabIndex        =   0
      Top             =   1050
      Width           =   3075
   End
   Begin Threed.SSCommand cmdReport 
      Height          =   375
      Left            =   10800
      TabIndex        =   8
      Top             =   1050
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "发放质保书"
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   6540
      Top             =   90
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Caption         =   "发放日期"
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
      Index           =   0
      Left            =   0
      Top             =   90
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "质量证明书编号"
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
      Left            =   0
      Top             =   555
      Width           =   1785
      _ExtentX        =   3149
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   6540
      Top             =   555
      Width           =   1020
      _ExtentX        =   1799
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
   Begin InDate.UDate dtp_fr_date 
      Height          =   315
      Left            =   7650
      TabIndex        =   9
      Tag             =   "发放日期"
      Top             =   90
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
   Begin InDate.UDate dtp_to_date 
      Height          =   315
      Left            =   9180
      TabIndex        =   10
      Tag             =   "发放日期"
      Top             =   90
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   4305
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10800
      Top             =   90
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "提货单号"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   7575
      Left            =   60
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   13361
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
      MaxCols         =   15
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQT0040C.frx":0000
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   10800
      Top             =   555
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Caption         =   "合同号"
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
   Begin Threed.SSCommand ssc_DIR_FIND 
      Height          =   345
      Left            =   9600
      TabIndex        =   12
      Top             =   1050
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      _Version        =   196609
      PictureFrames   =   1
      Picture         =   "AQT0040C.frx":0650
      ButtonStyle     =   1
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   315
      Left            =   1860
      TabIndex        =   13
      Top             =   1050
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption ssp_PRN 
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16448
         Caption         =   "直接打印"
      End
      Begin Threed.SSOption ssp_SAVE_PRN 
         Height          =   255
         Left            =   1290
         TabIndex        =   15
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   64
         Caption         =   "保存并打印"
      End
      Begin Threed.SSOption ssp_SAVE 
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12582912
         Caption         =   "保存不打印"
      End
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   0
      Top             =   1020
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "保存电子文档"
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
End
Attribute VB_Name = "AQT0040C"
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
'-- Program Name      质量证明书二次发放
'-- Program ID        AQD0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Chu Kyo Su
'-- Coder             Chu Kyo Su
'-- Date              2003.07. 25
'-- Description       质量证明书二次发放
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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim bPrintCheck As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'------------------------------ Report Variable ----------------------------------------------
'---------------------------------------------------------------------------------------------
Dim xlApp       As Object
Dim xlSheet     As Object

Dim arrRecords1 As Variant      'sQueryHeadC
Dim arrRecords2 As Variant      'sQueryDetailC
Dim arrRecords8 As Variant      'sQueryDetailC

Dim arrRecords3 As Variant      'sQueryHeadS
Dim arrRecords4 As Variant      'sQueryDetailS

Dim arrRecords5 As Variant      'sQueryHeadP
Dim arrRecords6 As Variant      'sQueryDetailP
Dim arrRecords7 As Variant      'sQueryDetailP

Dim arrRecords10 As Variant      'sQueryHeadb
Dim arrRecords11 As Variant      'sQueryDetailb

Dim sQuery      As String
Dim sErrMsg     As String
Dim sDate       As String
Dim AdoRs       As adodb.Recordset
Dim sPICTURE    As String
Dim oPICTURE    As Variant

'---------------------------------------------------------------------------------------------

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_CERT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PROD_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_fr_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_to_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_CUST_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ORD_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_SHP_ISP_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_PONO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
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
    
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AQD0030C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AQD0030C.P_REFER", Key:="P-R"
    sc1.Add Item:="AQD0030C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
     
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
            
        Case "txt_PROD_CD"             '产品
            sCode = "B0005"
                    
        Case "txt_CUST_CD"              '客户代码
            sCode = "CUST_CD"
            Set oCodeName = txt_CUST_NAME
            
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    Call subButtonHide

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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    txt_PROD_CD.Text = "PP"
    
    Screen.MousePointer = vbDefault
    
    Call subButtonHide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
    Call subButtonHide
    
End Sub



Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
     '  rControl(1).SetFocus
    End If
    txt_CERT_NO.Text = ""
    txt_CUST_CD.Text = ""
    txt_CUST_NAME.Text = ""
    txt_ORD_NO.Text = ""
    txt_PROD_CD.Text = ""
    txt_SHP_ISP_NO.Text = ""

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
     If subCheck = True Then
        
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                ss1.OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call subButtonHide
                Exit Sub
            End If
            
    Else
                
        GoTo Refer_Err
        
    End If
    
    Call subButtonHide
    
    bPrintCheck = False
    
    Exit Sub

Refer_Err:

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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim b_PRT_FL        As Boolean
Dim s_Msg           As String
Dim s_Shp_ist_no    As String

    If Row < 1 Then Exit Sub
    
    b_PRT_FL = Print_FL(Row)
    
    With ss1
        .Row = Row
        .Col = 10
        s_Shp_ist_no = .Text
        
        If b_PRT_FL Then
            .Col = 1
            If .Text = "1" Then
               .Col = 0:    .Text = "Update"
            Else
               .Col = 0:    .Text = ""
            End If
        Else
            .Col = 1
            If .Text = "1" Then
                .Col = 1: .Text = "0"
                .Col = 0: .Text = ""
                s_Msg = "提货单： " + s_Shp_ist_no + " 还有未回收的质保书,目前该提单质保书不可发放!"
                Call Gf_MessConfirm(s_Msg, "I")
            End If
        End If
    End With
End Sub



Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    If ss1.MaxRows > 0 And ss1.ActiveRow > 0 Then
        
        If Col <> 1 And Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 1) <> "" Then
            
            AQD0020C.Show
            AQD0020C.SetFocus
            AQD0020C.txt_CERT_NO.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 2)
            
            Call AQD0020C.Form_Ref
            
        End If
        
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
  '  If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 13)
   ' End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 13)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub


Private Sub subButtonHide()

    MDIMain.MenuTool.Buttons(4).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    

End Sub

'-----------------------------------------------------------------------
'---------------------------- Report Main ------------------------------
'-----------------------------------------------------------------------
Private Sub cmdReport_Click()
    Dim sCertNo As String
    Dim sFlag   As String
    Dim i       As Integer
    Dim Save_Path As String
    Dim Save_State As Integer
    Screen.MousePointer = vbHourglass
    
    Save_State = 0
    
        If ssp_PRN.Value = True Then
            Save_State = 0
        End If
        If ssp_SAVE_PRN.Value = True Then
            Save_State = 1
        End If
        If ssp_SAVE.Value = True Then
            Save_State = 2
        End If
    
    Save_Path = Trim(txt_SAVE_DIR)
    
    sErrMsg = ""
    
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    sDate = AdoRs.Fields(0)
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    With ss1
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Text = "1" Then
                .Col = 2:     sCertNo = Trim(.Text)
                .Col = 12:    sFlag = Trim(.Text)
                sErrMsg = Cert_type_check(sFlag, sCertNo, Save_State, Save_Path)
                                    
                If sErrMsg <> "" Then
                    i = .MaxRows
                End If
            End If
        Next i
        
        If sErrMsg = "" Then
            If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) = False Then Exit Sub
        End If
                            
    End With

    Call subButtonHide
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
        
End Sub


'--------------------------------------------------------------------------------------------------------
'------------------------------------------- Local Procedure --------------------------------------------
'--------------------------------------------------------------------------------------------------------

Private Function subCheck() As Boolean

    Dim sMesg As String
    Dim sFrDate As String
    Dim sToDate As String
    Dim sProdCd As String
    Dim sCertNo As String
    Dim sOrdNo As String
    Dim sTrnsNo As String
    Dim sCustCD As String
    
    sProdCd = Trim(txt_PROD_CD.Text)
    sCertNo = Trim(txt_CERT_NO.Text)
    sOrdNo = Trim(txt_ORD_NO.Text)
    sTrnsNo = Trim(txt_SHP_ISP_NO.Text)
    sCustCD = Trim(txt_CUST_CD.Text)
    
    sFrDate = Trim(dtp_fr_date.Text)
    sToDate = Trim(dtp_to_date.Text)
    
    sFrDate = Replace(sFrDate, "_", "")
    sToDate = Replace(sToDate, "_", "")
    
    sFrDate = Replace(sFrDate, "-", "")
    sToDate = Replace(sToDate, "-", "")
    
    If sCertNo = "" Then
        If sFrDate = "" Or sToDate = "" Then
            sMesg = "请完整输入发放日期（开始和结束日期）"
            Call Gp_MsgBoxDisplay(sMesg)
            subCheck = False
            Exit Function
        Else
            If sCustCD = "" And sOrdNo = "" And sProdCd = "" And sTrnsNo = "" Then
                sMesg = "请输入“产品代码”或“提货单号”或“客户代码”或“订单号”中的任意一项"
                Call Gp_MsgBoxDisplay(sMesg)
                subCheck = False
                Exit Function
            End If
        End If
    End If
         
    subCheck = True

End Function

Private Sub ssc_DIR_FIND_Click()
    Load Form_DIR_SELECT
    Form_DIR_SELECT.Show 1
   txt_SAVE_DIR.Text = sEXLSavePATH
End Sub

Private Sub txt_CERT_NO_Change()
    Call Gf_Control_text_Up(txt_CERT_NO)
End Sub

Private Sub txt_PROD_CD_Change()
    Call Gf_Control_text_Up(txt_PROD_CD)
End Sub

Private Function Print_FL(ByVal iRow As Long) As Boolean
     With ss1
     
          .Row = iRow
          .Col = 15
          If .Text = "n" Or .Text = "N" Then
             Print_FL = False
          Else
             Print_FL = True
          End If
                           
      End With
End Function

