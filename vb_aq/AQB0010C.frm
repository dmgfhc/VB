VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQB0010C 
   Caption         =   "质量设计接受现状查询 _AQB0010C"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_ORD_KND 
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
      Left            =   4005
      TabIndex        =   0
      Top             =   540
      Width           =   570
   End
   Begin VB.TextBox txt_PLT 
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
      Left            =   6360
      TabIndex        =   8
      Top             =   540
      Width           =   660
   End
   Begin VB.TextBox txt_Design_STS 
      Height          =   345
      Left            =   12840
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "0"
      Top             =   885
      Visible         =   0   'False
      Width           =   825
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8220
      Left            =   135
      TabIndex        =   7
      Top             =   930
      Width           =   15060
      _Version        =   393216
      _ExtentX        =   26564
      _ExtentY        =   14499
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   33
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQB0010C.frx":0000
   End
   Begin VB.TextBox txt_PROD_CD 
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
      Left            =   1125
      TabIndex        =   1
      Top             =   540
      Width           =   510
   End
   Begin VB.TextBox txt_PROD_NAME 
      Enabled         =   0   'False
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
      Left            =   1665
      TabIndex        =   2
      Top             =   540
      Width           =   1125
   End
   Begin VB.TextBox txt_CUST_CD 
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
      Left            =   6345
      TabIndex        =   3
      Top             =   135
      Width           =   1110
   End
   Begin VB.TextBox txt_CUST_NAME 
      Enabled         =   0   'False
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
      Left            =   7515
      TabIndex        =   4
      Top             =   135
      Width           =   5700
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   2
      Left            =   180
      Top             =   135
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      Caption         =   "订单确认日期"
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   0
      Left            =   180
      Top             =   540
      Width           =   885
      _ExtentX        =   1561
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   1
      Left            =   5400
      Top             =   135
      Width           =   885
      _ExtentX        =   1561
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
   Begin InDate.UDate dtp_CONF_DATE_MIN 
      Height          =   300
      Left            =   1665
      TabIndex        =   5
      Tag             =   "订单确认日期"
      Top             =   135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
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
   End
   Begin InDate.UDate dtp_CONF_DATE_MAX 
      Height          =   300
      Left            =   3195
      TabIndex        =   6
      Tag             =   "订单确认日期"
      Top             =   135
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
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
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   405
      Left            =   8520
      TabIndex        =   9
      Top             =   495
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   714
      _Version        =   196609
      BackStyle       =   1
      Begin VB.OptionButton opt_All_STS 
         BackColor       =   &H00E0E0E0&
         Caption         =   "全部的设计"
         Height          =   345
         Left            =   3420
         TabIndex        =   14
         Top             =   30
         Width           =   1215
      End
      Begin VB.OptionButton opt_Design_Error 
         BackColor       =   &H00E0E0E0&
         Caption         =   "设计出错"
         Height          =   345
         Left            =   2235
         TabIndex        =   12
         Top             =   30
         Width           =   1080
      End
      Begin VB.OptionButton opt_Confirm_Wait 
         BackColor       =   &H00E0E0E0&
         Caption         =   "等待确认"
         Height          =   345
         Left            =   1080
         TabIndex        =   11
         Top             =   30
         Width           =   1035
      End
      Begin VB.OptionButton opt_Confirm_End 
         BackColor       =   &H00E0E0E0&
         Caption         =   "已确认"
         Height          =   345
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   990
      End
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   3
      Left            =   7110
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "质量设计状态"
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   4
      Left            =   5400
      Top             =   540
      Width           =   885
      _ExtentX        =   1561
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   5
      Left            =   2880
      Top             =   540
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "订单类型"
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
   Begin Threed.SSCommand cmd_proc 
      Height          =   345
      Left            =   13290
      TabIndex        =   15
      Top             =   540
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
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
      Caption         =   "质量设计"
      BevelWidth      =   1
   End
End
Attribute VB_Name = "AQB0010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量设计
'-- Program Name      质量设计接受现状查询
'-- Program ID        AQB0010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       质量设计接受现状查询
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

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Sheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_PROD_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_PROD_NAME, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_CUST_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_CUST_NAME, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_CONF_DATE_MIN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_CONF_DATE_MAX, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_Design_STS, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_KND, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
      
      
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
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQB0010C.P_SREFER", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
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
            
        Case "txt_PROD_CD"              '产品
            sCode = "B0005"
            Set oCodeName = txt_PROD_NAME
            
        Case "txt_CUST_CD"              '客户代码
            sCode = "CUST_CD"
            Set oCodeName = txt_CUST_NAME
        Case "txt_ORD_KND"             '订单类型
            sCode = "B0009"
        Case "txt_PLT"                 '工厂
            sCode = "C0001"
                            
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub
'20080724 sunbin start
Private Sub cmd_proc_Click()

    Dim sMesg As String

   If Gf_Procedure_Exec("AQB1000P") = True Then

    '    sMesg = "Procedure Call is Success"

     '   Call Gp_MsgBoxDisplay(sMesg)

   End If
End Sub
'20080724 sunbin end

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call GP_MENU_SHOW_HIDE("09F")
    
    If dtp_CONF_DATE_MIN.RawData = "" Or dtp_CONF_DATE_MAX.RawData = "" Then
       dtp_CONF_DATE_MIN.Text = DateAdd("d", -1, Date)
       dtp_CONF_DATE_MAX.Text = Date
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
    If sAuthority = "1000" Then
       cmd_proc.Visible = False
    End If
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
'    Call GP_ROW_BACKCOLOR(ss1)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call GP_MENU_SHOW_HIDE("09F")
    
    Screen.MousePointer = vbDefault
    opt_All_STS.Value = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call GP_MENU_SHOW_HIDE("09F")
        txt_PROD_NAME.Text = ""
        txt_CUST_NAME.Text = ""
        dtp_CONF_DATE_MIN.Text = ""
        dtp_CONF_DATE_MAX.Text = ""
        pControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err
            
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call GP_MENU_SHOW_HIDE("09F")
        sOrderNo = Gf_Get_Cell_Value(ss1, ss1.Row, 1)
        sOrderItem = Gf_Get_Cell_Value(ss1, ss1.Row, 2)
        Call GP_SELECT_ROW(ss1, 1)
        
    End If
    
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_All_STS_Click()
    If opt_All_STS.Value = True Then
        txt_Design_STS.Text = "0"
    End If
End Sub

Private Sub opt_Confirm_End_Click()
    
    If opt_Confirm_End.Value = True Then
        txt_Design_STS.Text = "A"
    End If

End Sub

Private Sub opt_Confirm_Wait_Click()
    
    If opt_Confirm_Wait.Value = True Then
        txt_Design_STS.Text = "B"
    End If

End Sub

Private Sub opt_Design_Error_Click()
    
    If opt_Design_Error.Value = True Then
        txt_Design_STS.Text = "E"
    End If

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
    
    sOrderNo = Gf_GetCellText(ss1, ss1.ActiveRow, 1)
    sOrderItem = Gf_GetCellText(ss1, ss1.ActiveRow, 2)


End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    'Spread --> Control Value Move
    
    Dim iRow As Long
    
    With ss1

        iRow = Row

       If Row <> 0 Then
   
           Load AQB0020C
    
           .Row = Row
    
           .Col = 1: AQB0020C.txt_ORD_NO = .Text
           .Col = 2: AQB0020C.txt_ORD_ITEM = .Text
      
       End If
            
        AQB0020C.Show 1

   End With
      
   Call GP_SELECT_ROW(ss1, iRow)

    sOrderNo = Gf_Get_Cell_Value(ss1, iRow, 1)
    sOrderItem = Gf_Get_Cell_Value(ss1, iRow, 2)

End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    sOrderNo = Gf_Get_Cell_Value(ss1, NewRow, 1)
    sOrderItem = Gf_Get_Cell_Value(ss1, NewRow, 2)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub



Private Sub txt_CUST_CD_Change()
    If Trim(txt_CUST_CD.Text) = "" Then
        txt_CUST_NAME.Text = ""
    End If
End Sub



Private Sub txt_PROD_CD_Change()
    If Trim(txt_PROD_CD.Text) = "" Then
        txt_PROD_NAME.Text = ""
    End If
End Sub
