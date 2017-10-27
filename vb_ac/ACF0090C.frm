VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACF0090C 
   Caption         =   "生产成本数据跟踪明细_ACF0090C"
   ClientHeight    =   9225
   ClientLeft      =   285
   ClientTop       =   2325
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   12390
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_PLT 
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
      ItemData        =   "ACF0090C.frx":0000
      Left            =   4320
      List            =   "ACF0090C.frx":000D
      TabIndex        =   6
      Tag             =   "工厂代码"
      Text            =   "C1"
      Top             =   120
      Width           =   735
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8220
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   14499
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACF0090C.frx":001D
      Begin FPSpread.vaSpread ss1 
         Height          =   2340
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   19500
         _Version        =   393216
         _ExtentX        =   34396
         _ExtentY        =   4128
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
         MaxCols         =   43
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0090C.frx":00AF
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   2160
         Left            =   0
         TabIndex        =   2
         Top             =   3675
         Width           =   19500
         _Version        =   393216
         _ExtentX        =   34396
         _ExtentY        =   3810
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
         MaxCols         =   11
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0090C.frx":1768
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   1245
         Left            =   0
         TabIndex        =   3
         Top             =   2385
         Width           =   19500
         _Version        =   393216
         _ExtentX        =   34396
         _ExtentY        =   2196
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
         MaxCols         =   10
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0090C.frx":1F8B
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   2340
         Left            =   0
         TabIndex        =   4
         Top             =   5880
         Width           =   19500
         _Version        =   393216
         _ExtentX        =   34396
         _ExtentY        =   4128
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
         MaxCols         =   53
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0090C.frx":2692
      End
   End
   Begin InDate.ULabel ULabel3 
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
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
   Begin InDate.UDate prod_date_from 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Tag             =   "开始日期"
      Top             =   120
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
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3120
      Top             =   120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "工厂代码"
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
   Begin Threed.SSCommand Cmd_Edit 
      Height          =   360
      Left            =   7680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "更新数据"
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   15120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   15105
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "ACF0090C"
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
'-- Program ID        ACB1022C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2003.9.26
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
Dim Sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS2_PLAN_DELAY = 3  '计划
Const SS2_REPAIR_DELAY = 4           '定修
Const SS2_MACH_DELAY = 5   '机械
Const SS2_ELECT_DELAY = 6            '电器
Const SS2_OPER_DELAY = 7 '操作
Const SS2_NON_PLAN_DELAY = 8    '故障


'Const SS2_PLT = 1




Dim sWgtLenFlag As String
Dim sQuery  As String

Private Sub Form_Define()

 Dim iRow As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     'Call Gp_Ms_Collection(prod_date_from, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(prod_date_from, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(CBO_PLT, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
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
     Call Gp_Sp_Collection(ss1, 1, "", " ", " ", "", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

      For iRow = 5 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
      Next iRow
    

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACF0090C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    
'    sc1.Item("Spread").Col = 0
'    sc1.Item("Spread").Row = 0
'    sc1.Item("Spread").Text = "◎"


    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     For iRow = 5 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iRow, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
     Next iRow
    

    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACF0090C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    'Call Gp_Sp_ColHidden(ss2, 13, True)
    
    
    
'    sc2.Item("Spread").Col = 0
'    sc2.Item("Spread").Row = 0
'    sc2.Item("Spread").Text = "◎"

     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     
    For iRow = 5 To ss3.MaxCols
        Call Gp_Sp_Collection(ss3, iRow, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3, True)
     Next iRow


    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="ACF0090C.P_REFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    
    
     Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     
     For iRow = 5 To ss4.MaxCols
        Call Gp_Sp_Collection(ss4, iRow, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4, True)
     Next iRow
    

    
    'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="ACF0090C.P_REFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc4, Key:="Sc4"
    
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    
    
    
'    Sc3.Item("Spread").Col = 0
'    Sc3.Item("Spread").Row = 0
'    Sc3.Item("Spread").Text = "◎"
    
        
End Sub


Private Sub Cmd_Edit_Click()
  Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
          
    If Trim(prod_date_from) = "" Then
        Call Gp_MsgBoxDisplay(prod_date_from.Tag + "必须输入")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACF0090P ('" + Trim(Format(prod_date_from.Text, "YYYYMMDD")) + "',?)}"

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
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("更新成功..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste

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
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(sc4.Item("Spread"), False)

    
    'Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
   ' Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
'    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(sc4)

    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "C-System.INI", Me.Name)

    
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 5)
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 7)
    'Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 14)
    'Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 15)
    
    Screen.MousePointer = vbDefault
    
    
    prod_date_from.Text = Date - 1
    
    
    CBO_PLT.Text = "C1"
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "C-System.INI", Me.Name)

    
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
    Set Sc3 = Nothing
    Set sc4 = Nothing

    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) And Gf_Sp_Cls(Sc3) And Gf_Sp_Cls(sc1) And Gf_Sp_Cls(sc4) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rContro1(1).SetFocus
            CBO_PLT.Text = "C1"
    End If
    
End Sub

Public Sub Form_Ref()

         Call Gf_Sp_Cls(sc2)
         Call Gf_Sp_Cls(Sc3)
         Call Gf_Sp_Cls(sc1)
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rContro1(1).SetFocus
   
       
    

    
    'If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
        ss1.OperationMode = OperationModeNormal
        'Call Gp_Sp_BlockColor(ss1, 7, 7, 1, ss1.MaxRows)
        Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss2.OperationMode = OperationModeNormal
        
        Call Gf_Sp_Cls(Sc3)
        Call Gf_Sp_Refer(M_CN1, Sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss3.OperationMode = OperationModeNormal
        
        Call Gf_Sp_Cls(sc4)
        Call Gf_Sp_Refer(M_CN1, sc4, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss4.OperationMode = OperationModeNormal
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        MDIMain.MenuTool.Buttons(4).Enabled = True
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
                  
        
End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Form_Exc()


'    If txt_shape.Text = "ss1" Then
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'    ElseIf txt_shape.Text = "ss2" Then
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'    ElseIf txt_shape.Text = "ss3" Then
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'     ElseIf txt_shape.Text = "ss4" Then
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'     ElseIf txt_shape.Text = "ss5" Then
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'     ElseIf txt_shape.Text = "ss6" Then
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'     ElseIf txt_shape.Text = "ss7" Then
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'    Else
'        Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
'    End If

   Dim i               As Integer
    Dim j               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    Dim sShift          As String
    
    Dim sPage_Num       As Integer
    Dim sPage_X         As Integer
    Dim sPage           As Double
    Dim sLastPage       As Double
    Dim sRow1           As Integer
    Dim sRow2           As Integer
    
    Dim xl_1            As String
    Dim xl_2            As String
    Dim xl_3            As String
    Dim xl_4            As String
    Dim xl_5            As String
    Dim xl_6            As String
    Dim xl_7            As String
    'Dim xl_H            As String
    'Dim xl_I            As String
'    Dim xl_J            As String
    
    Dim xl_clr_body     As String
    Dim xl_clr_sum      As String
    Dim xl_clr_spc      As String
    
    Dim Xl_Cnt          As String
    Dim Xl_Wgt          As String
    Dim Xl_Wgt_Val      As String
    Dim Xl_Ust          As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If ERR.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ERR.Clear

    xlApp.Workbooks.Open (App.Path & "\ACF0090C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = prod_date_from.Text
    
        xlApp.Range("B4").Value = Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日"
    
    

    xlApp.Range("D90").Value = Now
    
    xlApp.Range("P90").Value = sUserName
        
xlApp.Application.Visible = True
    

        xl_1 = "C5:R9"
        xl_2 = "C13:L17"
        xl_3 = "C21:R25"
        xl_4 = "C28:C36"
        xl_5 = "C40:L44"
        xl_6 = "C49:AF53"
        xl_7 = "C57:X61"
        
        Clipboard.Clear
        ss1.SetSelection 2, 1, 17, 5
        ss1.ClipboardCopy
        xlApp.Range(xl_1).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss1.ClearSelection
        Sleep 100
        
        Clipboard.Clear
        ss1.SetSelection 18, 1, 27, 5
        ss1.ClipboardCopy
        xlApp.Range(xl_2).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss1.ClearSelection
        Sleep 100
        
        Clipboard.Clear
        ss1.SetSelection 28, 1, 43, 5
        ss1.ClipboardCopy
        xlApp.Range(xl_3).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss1.ClearSelection
        Sleep 100
        
        
'        Clipboard.Clear
'        ss2.SetSelection 3, 1, ss2.MaxCols, ss2.MaxRows
'        ss2.ClipboardCopy
'        xlApp.Range(xl_4).Select
'        xlApp.ActiveSheet.Paste
'        Clipboard.Clear
'        ss2.ClearSelection
'        Sleep 100
        
        Clipboard.Clear
        ss3.SetSelection 2, 1, 11, 5
        ss3.ClipboardCopy
        xlApp.Range(xl_5).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss3.ClearSelection
        
        Sleep 100
        
        Clipboard.Clear
        ss4.SetSelection 2, 1, 31, 5
        ss4.ClipboardCopy
        xlApp.Range(xl_6).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss4.ClearSelection
        Sleep 100
        
        Clipboard.Clear
        ss4.SetSelection 32, 1, 53, 5
        ss4.ClipboardCopy
        xlApp.Range(xl_7).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        ss4.ClearSelection
        Sleep 100
        
        
    ss2.Row = 1
    ss2.Col = 2: xlApp.Range("C28").Value = ss2.Text
    ss2.Col = 3: xlApp.Range("C29").Value = ss2.Text
    ss2.Col = 4: xlApp.Range("C30").Value = ss2.Text
    ss2.Col = 5: xlApp.Range("C31").Value = ss2.Text
    ss2.Col = 6: xlApp.Range("C32").Value = ss2.Text
    ss2.Col = 7: xlApp.Range("C33").Value = ss2.Text
    ss2.Col = 8: xlApp.Range("C34").Value = ss2.Text
    ss2.Col = 9: xlApp.Range("C35").Value = ss2.Text
    ss2.Col = 10: xlApp.Range("C36").Value = ss2.Text
    
    If CBO_PLT.Text = "C1" Then
    
    xlApp.Range("B1").Value = "中厚板卷厂"
    
    ElseIf CBO_PLT.Text = "C2" Then
    
    xlApp.Range("B1").Value = "宽厚板厂"
    
    ElseIf CBO_PLT.Text = "C3" Then
    
    xlApp.Range("B1").Value = "中板厂"
    
    End If
    
        
'
'    ss1.ClearSelection
'    ss2.ClearSelection
'    ss3.ClearSelection
'    ss4.ClearSelection
'    ss5.ClearSelection
'    ss6.ClearSelection
'    ss7.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    'xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault


End Sub



'
Public Sub Form_Exit()
    Unload Me
End Sub

