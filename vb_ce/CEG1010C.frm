VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form CEG1010C 
   Caption         =   "月坯料使用计划对象订单查询/选定_CEG1010C"
   ClientHeight    =   7905
   ClientLeft      =   75
   ClientTop       =   2355
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   15165
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_ord_knd 
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
      Left            =   5985
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "订单种类"
      Top             =   120
      Width           =   345
   End
   Begin VB.TextBox txt_ord_knd_nm 
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
      Left            =   6330
      TabIndex        =   3
      Tag             =   "订单种类"
      Top             =   120
      Width           =   1365
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   1365
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   1740
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   120
      Width           =   1815
   End
   Begin InDate.UDate txt_del_fr 
      Height          =   315
      Left            =   1365
      TabIndex        =   4
      Tag             =   "交货期"
      Top             =   510
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   180
      Top             =   510
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "交货期"
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
   Begin Threed.SSCommand cmd_slab_confirm 
      Height          =   405
      Left            =   10980
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "坯料选定处理"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_fur_balance 
      Height          =   405
      Left            =   12420
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "加热炉均衡"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_plt_balance 
      Height          =   405
      Left            =   13860
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "炼钢厂均衡"
      BevelWidth      =   3
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8205
      Left            =   60
      TabIndex        =   11
      Top             =   990
      Width           =   15210
      _Version        =   393216
      _ExtentX        =   26829
      _ExtentY        =   14473
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
      MaxCols         =   35
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CEG1010C.frx":0000
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   180
      Top             =   120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "工 厂"
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
   Begin InDate.UDate txt_del_to 
      Height          =   315
      Left            =   2835
      TabIndex        =   5
      Tag             =   "交货期"
      Top             =   510
      Width           =   1440
      _ExtentX        =   2540
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
   Begin Threed.SSCommand cmd_slab_design_change 
      Height          =   405
      Left            =   9540
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "设计标准变更"
      BevelWidth      =   3
   End
   Begin InDate.UDate txt_release_fr 
      Height          =   315
      Left            =   5985
      TabIndex        =   6
      Tag             =   "交货期"
      Top             =   510
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.UDate txt_release_to 
      Height          =   315
      Left            =   7470
      TabIndex        =   7
      Tag             =   "交货期"
      Top             =   510
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   4800
      Top             =   510
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "投入生产期"
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   4800
      Top             =   120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "订单种类"
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
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   75
      X2              =   15240
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   15240
      Y1              =   945
      Y2              =   945
   End
End
Attribute VB_Name = "CEG1010C"
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
'-- Program Name      SLAB USE MONTHLY PLAN
'-- Program ID        CEG1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2007.10.22
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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Active_Row As Integer           'Active Row

Private Sub Form_Define()

    Dim iCol As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_plt, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_del_fr, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_del_to, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_release_fr, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_release_to, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ord_knd, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_ord_knd_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 6, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 7, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 8, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    
    For iCol = 9 To 21
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    For iCol = 23 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CEG1010C.P_REFER1", Key:="P-R"
    sc1.Add Item:="CEG1010C.P_ONEROW1", Key:="P-O"
    sc1.Add Item:="CEG1010C.P_MODIFY1", Key:="P-M"
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
    
    Call Gp_Sp_ColHidden(ss1, 27, True)
    Call Gp_Sp_ColHidden(ss1, 28, True)
    Call Gp_Sp_ColHidden(ss1, 29, True)
    Call Gp_Sp_ColHidden(ss1, 30, True)
    
End Sub

Private Sub cmd_fur_balance_Click()
    CEG1040C.Show
    CEG1040C.SetFocus
End Sub

Private Sub cmd_plt_balance_Click()
    CEG1050C.Show
    CEG1050C.SetFocus
End Sub

Private Sub cmd_slab_confirm_Click()

    On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    ret_Result_ErrMsg = Gf_Ms_NeceCheck(Mc1("nControl"))
    If ret_Result_ErrMsg <> "OK" Then
        ret_Result_ErrMsg = ret_Result_ErrMsg + "必须输入"
        Call Gp_MsgBoxDisplay(ret_Result_ErrMsg, "", "错误提示")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call CEG1020P ('" & txt_plt.Text & "','" & txt_del_fr.RawData & "','" & txt_del_to.RawData & _
                                                  "','" & txt_release_fr.RawData & "','" & txt_release_to.RawData & _
                                                  "','" & txt_ord_knd.Text & "','" & sUserID & "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
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
        Call Gp_MsgBoxDisplay("坯料选定完了..!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, Err.Description
    
End Sub

Private Sub cmd_slab_design_change_Click()

    Load Slab_Design_Change
    
    If Active_Row > 0 Then
    
        ss1.Row = Active_Row
        
        ss1.Col = 1
        Slab_Design_Change.txt_ord_no.Text = ss1.Text
        ss1.Col = 2
        Slab_Design_Change.txt_ORD_ITEM.Text = ss1.Text
        ss1.Col = 3
        Slab_Design_Change.txt_stlgrd.Text = ss1.Text
        ss1.Col = 4
        Slab_Design_Change.txt_stlgrd_name.Text = ss1.Text
        ss1.Col = 6
        Slab_Design_Change.sdb_slab_thk_fr.Value = ss1.Value
        Slab_Design_Change.sdb_slab_thk_to.Value = ss1.Value
        ss1.Col = 7
        Slab_Design_Change.sdb_slab_wid_fr.Value = ss1.Value
        Slab_Design_Change.sdb_slab_wid_to.Value = ss1.Value
        ss1.Col = 8
        Slab_Design_Change.sdb_slab_len_fr.Value = 0
        Slab_Design_Change.sdb_slab_len_to.Value = 9999999
        ss1.Col = 13
        Slab_Design_Change.sdb_asroll_cnt_fr.Value = 0
        Slab_Design_Change.sdb_asroll_cnt_to.Value = 99
        ss1.Col = 14
        Slab_Design_Change.sdb_prod_thk_fr.Value = 0
        Slab_Design_Change.sdb_prod_thk_to.Value = 9999.99
        ss1.Col = 15
        Slab_Design_Change.sdb_prod_wid_fr.Value = 0
        Slab_Design_Change.sdb_prod_wid_to.Value = 9999.99
        ss1.Col = 16
        Slab_Design_Change.sdb_prod_len_fr.Value = 0
        Slab_Design_Change.sdb_prod_len_to.Value = 9999999.9
        
        Slab_Design_Change.cbo_prod_cnt.Text = "0"
        
        Slab_Design_Change.txt_fur_line.Text = "1"
        Slab_Design_Change.opt_fur_no1.Value = True
        
    Else
    
        Slab_Design_Change.sdb_slab_thk_to.Value = 9999
        Slab_Design_Change.sdb_slab_wid_to.Value = 9999
        Slab_Design_Change.sdb_slab_len_to.Value = 9999999
        Slab_Design_Change.sdb_asroll_cnt_to.Value = 99
        Slab_Design_Change.sdb_prod_thk_to.Value = 9999.99
        Slab_Design_Change.sdb_prod_wid_to.Value = 9999.99
        Slab_Design_Change.sdb_prod_len_to.Value = 9999999.9
        
        Slab_Design_Change.cbo_prod_cnt.Text = "0"
        
        Slab_Design_Change.txt_fur_line.Text = "1"
        Slab_Design_Change.opt_fur_no1.Value = True
        
    End If
    
    Slab_Design_Change.P_MODE = "M"
    
    Slab_Design_Change.Show 1
    
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet
   
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
    
    If Mid(sAuthority, 3, 1) <> "1" Then
        cmd_slab_design_change.Enabled = False
        cmd_slab_confirm.Enabled = False
    End If
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    
    txt_plt.Text = "C3"
    Call txt_plt_KeyUp(0, 0)
    
    txt_del_fr.Text = Mid(txt_del_fr.Text, 1, 8) + "01"
    
    txt_del_to.Text = Format(DateAdd("m", 1, txt_del_fr.Text), "YYYY-MM-DD")
    txt_del_to.Text = DateAdd("d", -1, txt_del_to.Text)
    
    txt_release_fr.Text = Mid(DateAdd("M", -1, txt_release_fr.Text), 1, 8) & "20"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    
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
    Call MenuTool_ReSet
    
End Sub

Public Sub Spread_Can()
    
    Call Gp_Sp_Cancel(M_CN1, sc1)
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        txt_plt.Text = "C3"
        Call txt_plt_KeyUp(0, 0)
        Active_Row = 0
        txt_del_fr.Text = Mid(txt_del_fr.Text, 1, 8) + "01"
        txt_del_to.Text = Format(DateAdd("m", 1, txt_del_fr.Text), "YYYY-MM-DD")
        txt_del_to.Text = DateAdd("d", -1, txt_del_to.Text)
        txt_release_fr.Text = Mid(DateAdd("M", -1, txt_release_fr.Text), 1, 8) & "20"
    End If
        
End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    Active_Row = 0
    
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
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
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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
    
    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Active_Row = Row
    
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    ss1.Col = 5
    ss1.Row = Row
    If ss1.Text <> "N" Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If

End Sub

Private Sub txt_ord_knd_DblClick()

    Call txt_ord_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_ord_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0009"

        DD.rControl.Add Item:=txt_ord_knd
        DD.rControl.Add Item:=txt_ord_knd_nm

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_ord_knd.Text)) = txt_ord_knd.MaxLength Then
            txt_ord_knd_nm.Text = Gf_ComnNameFind(M_CN1, "B0009", txt_ord_knd.Text, 2)
            Exit Sub
        Else
            txt_ord_knd_nm.Text = ""
        End If
        
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

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub
