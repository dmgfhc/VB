VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1300C 
   Caption         =   "编制剩余生产计划_AAA1300C"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_excel 
      Height          =   270
      Left            =   2850
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   195
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txt_ord_yn 
      Height          =   270
      Left            =   8250
      MaxLength       =   1
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.ComboBox cbo_plt 
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
      ItemData        =   "AAA1300C.frx":0000
      Left            =   4545
      List            =   "AAA1300C.frx":0002
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   105
      Width           =   870
   End
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
      ItemData        =   "AAA1300C.frx":0004
      Left            =   7245
      List            =   "AAA1300C.frx":0011
      TabIndex        =   2
      Tag             =   "产品"
      Top             =   105
      Width           =   870
   End
   Begin Threed.SSCommand order_cmd 
      Height          =   330
      Left            =   12375
      TabIndex        =   3
      Top             =   105
      Width           =   1305
      _ExtentX        =   2302
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
      Caption         =   "订单查询"
   End
   Begin InDate.ULabel ULabel5 
      Height          =   300
      Left            =   5895
      Top             =   105
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
   Begin InDate.ULabel ULabel3 
      Height          =   300
      Left            =   3195
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
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
   End
   Begin InDate.UDate dtp_yy_mm 
      Height          =   300
      Left            =   1470
      TabIndex        =   0
      Tag             =   "年月"
      Top             =   105
      Width           =   1185
      _ExtentX        =   2090
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
      Left            =   120
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "年月"
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
   Begin Threed.SSCommand plan_cmd 
      Height          =   330
      Left            =   13770
      TabIndex        =   4
      Top             =   105
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
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
      Caption         =   "编制计划"
   End
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Left            =   8565
      Top             =   105
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      Caption         =   "基准时间:"
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
      ForeColor       =   -2147483641
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8730
      Left            =   105
      TabIndex        =   6
      Top             =   480
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   15399
      _Version        =   196609
      PaneTree        =   "AAA1300C.frx":0021
      Begin FPSpread.vaSpread ss1 
         Height          =   855
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   1508
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
         MaxCols         =   20
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         SpreadDesigner  =   "AAA1300C.frx":0093
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   2175
         Left            =   30
         TabIndex        =   8
         Top             =   975
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   3836
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
         SpreadDesigner  =   "AAA1300C.frx":0D0C
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   5460
         Left            =   30
         TabIndex        =   9
         Top             =   3240
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   9631
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
         SpreadDesigner  =   "AAA1300C.frx":0F32
      End
   End
End
Attribute VB_Name = "AAA1300C"
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
'-- Program ID        AAA1300C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2009.5.4
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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
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
      Call Gp_Ms_Collection(dtp_yy_mm, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ord_yn, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '预测产能
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AAA1300C.P_SREFER1", Key:="P-R"
    Sc1.Add Item:="AAA1300C.P_SMODIFY1", Key:="P-M"
    
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"

    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    
    Call Gp_Sp_ColHidden(ss1, 1, True)
    Call Gp_Sp_ColHidden(ss1, 2, True)
    
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AAA1300C.P_SREFER3", Key:="P-R"
    sc3.Add Item:="AAA1300C.P_SMODIFY3", Key:="P-M"
    
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    
    Call Gp_Sp_ColHidden(ss3, 1, True)
    Call Gp_Sp_ColHidden(ss3, 2, True)
    Call Gp_Sp_ColHidden(ss3, 3, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub order_cmd_Click()
   
   txt_ord_yn = "1"
   Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc1, Mc1("nControl"))
   Call Sp_Header_Refer
''       Call Gp_Sp_ColLock(ss3, 13, True)

End Sub

Private Sub plan_cmd_Click()

On Error GoTo plan_cmd_Error

    Dim sQuery As String
    Dim iCount As Integer
    
    'If dtp_date_str.Enabled Then Exit Sub
    
    Dim adoCmd As ADODB.Command
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    For iCount = 1 To 2
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount), adVariant, adParamOutput)
    Next iCount
    
    'CAST
    sQuery = "{call AAA5010P ('" + dtp_yy_mm.RawData + "', '" + cbo_plt.Text + "',  '" + sUserID + "', ?,? )}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(1) <> "" Then
        Call Gp_MsgBoxDisplay(adoCmd(1))
        Set adoCmd = Nothing
        Exit Sub
    End If
        
    Set adoCmd = Nothing
    
    Call Form_Ref
    
    Exit Sub

plan_cmd_Error:

    Call Gp_MsgBoxDisplay("编制计划错误 : " & Error)

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

    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    Call Sp_Setting2(ss1)
    Call Sp_Setting2(ss2)
    
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    txt_prod_cd.Text = "PP"
    
    cbo_plt.AddItem "C1"
    cbo_plt.AddItem "C3"
    
    Call Gp_Sp_ColGet(ss1, "A-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "A-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "A-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
    If Mid(sAuthority, 1, 3) = "111" Then
       order_cmd.Enabled = True
       plan_cmd.Enabled = True
    Else
       order_cmd.Enabled = False
       plan_cmd.Enabled = False
    End If

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
    
    Set pColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set iColumn1 = Nothing
    Set aColumn1 = Nothing
    Set lColumn1 = Nothing
    
    Set pColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set iColumn3 = Nothing
    Set aColumn3 = Nothing
    Set lColumn3 = Nothing
    
    Set THK_GRP = Nothing
    Set WID_GRP = Nothing
    Set MIN_VALUE = Nothing
    Set MAX_VALUE = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call Gp_Sp_ColSet(ss1, "A-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss2, "A-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss3, "A-System.INI", Me.Name)
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    rControl(1).SetFocus
    ULabel2.Caption = "基准时间: "
    
    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, False
    'ss2.ClearRange 1, 1, ss2.MaxCols, ss2.MaxRows, False
    'ss3.ClearRange 1, 1, ss3.MaxCols, ss3.MaxRows, False
    ss2.MaxRows = 0
    ss2.MaxCols = 0
    ss3.MaxRows = 0
End Sub

Public Sub Form_Ref()
    txt_ord_yn = ""
    'Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"))
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl")) Then
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc1) Then
           Call Sp_Header_Refer
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    '       Call Gp_Sp_ColLock(ss3, 13, False)
           txt_excel = "1"
           
           ss1.Row = 1
           ss1.Col = 16
           ULabel2.Caption = "基准时间 ：  " + ss1.Text + "日 "
           ss1.Col = 17
           ULabel2.Caption = ULabel2.Caption + ss1.Text + "时"
        End If
    End If
    

End Sub

Public Sub Form_Pro()
    
    Dim iRow As Integer
    If txt_ord_yn = "1" Then
    
        Call old_data_delete
        
        With ss3
            .Col = 0
            For iRow = 1 To .MaxRows
            .Row = iRow
            .Text = "Update"
            Next iRow
        End With
     End If
    
    'If Sp_Process(M_CN1, Proc_Sc("Sc1")) And Sp_Process(M_CN1, Proc_Sc("Sc3")) Then
    
'    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc1"), Nothing, True) Or Gf_Sp_Process(M_CN1, Proc_Sc("Sc3"), Nothing, True) Then
    
    Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc1"), Nothing, True)
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc3"), Nothing, True) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Form_Ref
    End If
    
End Sub

Public Sub Form_Exc()
If txt_excel.Text = "1" Then
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
ElseIf txt_excel.Text = "2" Then
    Call Pp_Sp_Excel(Me, ss2, 0, lBlkcol2, lBlkrow1, lBlkrow2)
ElseIf txt_excel.Text = "3" Then
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If
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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
txt_excel = "1"
End Sub

Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)

Dim Rem_Temper, Ave_Th, Ave_Wrk_Rate As Double
Dim Pln_work_time, Pln_pure_time As Double
Dim Act_pure_time, Ave_pure_time As Double

If ss1.ActiveCol = 13 Then
   ss1.Row = ss1.ActiveRow
   ss1.Col = ss1.ActiveCol
   Ave_Wrk_Rate = Val(ss1.Text)
   
   ss1.Col = 5
   Pln_work_time = Val(ss1.Text)
   ss1.Col = 6
   Ave_pure_time = Val(ss1.Text)
   ss1.Col = 7
   Act_pure_time = Val(ss1.Text)
   Pln_pure_time = Round(Pln_work_time * (Ave_Wrk_Rate / 100), 3)
   Rem_Temper = Pln_pure_time - Act_pure_time
   
'   ss1.Col = 10
'   Rem_Temper = Val(ss1.Text)
   ss1.Col = 11
   Ave_Th = Val(ss1.Text)
   ss1.Col = 14
   ss1.Text = Rem_Temper * Ave_Th  '* Ave_Wrk_Rate / 100
   If dtp_yy_mm.Enabled = False Then
      ss1.Col = 0
      ss1.Text = "Update"
   End If
End If
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
txt_excel = "2"
End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
txt_excel = "3"
End Sub

Private Sub ss3_EditChange(ByVal Col As Long, ByVal Row As Long)
      ss3.Row = ss3.ActiveRow
      ss3.Col = 0
      ss3.Text = "Update"
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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
    Dim sQuery3 As String
    
    Dim AdoRs2 As ADODB.Recordset
    Dim ArrayRecords2 As Variant
    
    Dim AdoRs3 As ADODB.Recordset
    Dim ArrayRecords3 As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "SELECT THK_CD, FR_THK, TO_THK "
    sQuery = sQuery + "   FROM BP_THICK_GRP "
    sQuery = sQuery + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery = sQuery + "    AND THK_CD <> '*' "
    sQuery = sQuery + "  ORDER BY THK_CD "
    
    With ss2

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
        
            .MaxCols = UBound(ArrayRecords, 2) + 1
        
            For iCol = 0 To .MaxCols - 1
            
               .Col = iCol + 1
               .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                           
                'Column Type Setting
                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 3
                .TypeNumberMax = 99999999999.999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                .ColWidth(iCol + 1) = 13
                                
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
    
    With ss2

        Sp_Header_Refer = True
        .ReDraw = False
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
                                
                .Row = iRow: .Row2 = iRow
                .Col = 1: .Col2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 3
                .TypeNumberMax = 99999999999.999
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

    Set AdoRs3 = New ADODB.Recordset
    
    sQuery3 = "{ call AAA1300C.P_SREFER2('" + dtp_yy_mm.RawData + "','" + cbo_plt.Text + "', '" + txt_prod_cd.Text + "', '" + txt_ord_yn.Text + "') }"
    
    AdoRs3.Open sQuery3, M_CN1, adOpenKeyset
        
    If AdoRs3.BOF Or AdoRs3.EOF Then
    
        Sp_Header_Refer = False
        AdoRs3.Close
        Set AdoRs3 = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If
        
    ArrayRecords3 = AdoRs3.GetRows
    
    AdoRs3.Close
    Set AdoRs3 = Nothing
    
    With ss2
        iCnt = 0
        For iCnt = 0 To UBound(ArrayRecords3, 2)
            .Row = Asc(ArrayRecords3(0, iCnt)) - 64

            For iCol = 1 To .MaxCols
                .Col = iCol
                
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 3
                .TypeNumberMax = 99999999999.999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
    
                If ArrayRecords3(iCol, iCnt) = vbNull Or ArrayRecords3(iCol, iCnt) = 0 Then
                    .Text = ""
                Else
                    .Text = CStr(ArrayRecords3(iCol, iCnt))
                End If
            Next iCol
        Next iCnt
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Set AdoRs2 = Nothing
    Set AdoRs3 = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     
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

Public Sub Sp_Setting2(ByVal sPname As Variant)

    With sPname
    
        .RowHeight(-1) = 12
        .RowHeight(0) = 16
        
'        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .OperationMode = OperationModeNormal
        '.RetainSelBlock = True

        '.UserResize = UserResizeNone
        
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
        
'        .Col = 0
'        .Row = -1
'        .FontBold = True
        
        .LockBackColor = RGB(255, 255, 255)
        
'        If .Name = "ss3" Then Call Gp_Sp_RowColor(ss3, 3, vbRed)
'        If .Name = "ss4" Then .RowHeadersShow = False
        
    End With
    
End Sub

Public Sub Pp_Sp_Excel(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
'        If bLkcol1 = 0 Then
'           bLkcol1 = 1
'        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .Row = bLkrow1: .Row2 = bLkrow2
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        xlSheet.Range("A1").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .Row = 1
            
            iExlCol = ColIndex
            If IsNumeric(.Text) And Left(.Text, 1) = "0" And _
               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
                If ColIndex > 104 Then
                    sExlRange1 = "D" & sExlRange1
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C" & sExlRange1
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B" & sExlRange1
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If
                
                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000"
                ElseIf Len(.Text) = 10 Then
                    xlSheet.Range(sExlRange).NumberFormat = "0000000000"
                ElseIf Len(.Text) = 12 Then
                    xlSheet.Range(sExlRange).NumberFormat = "000000000000"
                ElseIf Len(.Text) = 14 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000000000"
                End If
            End If
        Next
    
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel", "W")

End Sub
Private Sub old_data_delete()

On Error GoTo old_data_Error

    Dim sQuery As String
    Dim iCount As Integer
  
    Dim adoCmd As ADODB.Command
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    For iCount = 1 To 2
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount), adVariant, adParamOutput)
    Next iCount
    
    ' --txt_ord_yn
    sQuery = "{call AAA1300C.P_DELETE3 ('" + dtp_yy_mm.RawData + "', '" + cbo_plt.Text + "', '" + txt_prod_cd.Text + "', ?,? )}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(1) <> "" Then
        Call Gp_MsgBoxDisplay(adoCmd(1))
        Set adoCmd = Nothing
        Exit Sub
    End If
        
    Set adoCmd = Nothing
    
    Exit Sub

old_data_Error:

    Call Gp_MsgBoxDisplay("旧的信息删掉 错误 : " & Error)

End Sub


