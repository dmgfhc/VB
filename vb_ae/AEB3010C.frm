VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AEB3010C 
   Caption         =   "相关母板/产品信息_AEB3010C"
   ClientHeight    =   6450
   ClientLeft      =   405
   ClientTop       =   3915
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   14355
   Begin CSTextLibCtl.sidbEdit TXT_SLAB_NO 
      Height          =   315
      Left            =   1935
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1095
      _Version        =   262145
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
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
      Enabled         =   0   'False
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
      StartText.x     =   2
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   5
      Undo            =   0
      Data            =   0
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   6315
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   14235
      _Version        =   393216
      _ExtentX        =   25109
      _ExtentY        =   11139
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
      MaxCols         =   16
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AEB3010C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   135
      Top             =   180
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      Caption         =   "板坯编制号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CMD1 
      Height          =   420
      Left            =   13140
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
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
      Caption         =   "退出"
   End
   Begin Threed.SSCommand cmd_del 
      Height          =   420
      Left            =   12060
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      MarqueeDirection=   1
      ForeColor       =   255
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
      Caption         =   "删除"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   45
      X2              =   14260
      Y1              =   495
      Y2              =   495
   End
End
Attribute VB_Name = "AEB3010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       工序计划
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEB3010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          jianing
'-- Coder             jianing
'-- Date              2003.6.19
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType   As String           'Form Type
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

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(TXT_SLAB_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
         
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    'Duplicate Count
    'iDupCnt = 1
    
    'Sum Column Count
    'iSumCnt = 1
    
    'Sum Column Setting
    'iSumCol.Add Item:=4
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Me.Top = 3080
    Me.Left = 620

    Call Gp_Sp_ColHidden(ss1, 9, True)

End Sub

Private Sub cmd_del_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    
    Dim adoCmd As adodb.Command
    
    If ss1.MaxRows = 0 Then Exit Sub
    
    If Not Gf_MessConfirm("Do you erase data really ?", "Q") Then Exit Sub
    
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
    
    sQuery = "{call AEB2060C.P_MODIFY ('S'," & TXT_SLAB_NO.Value & ",0,0,?,?)}"
    
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
    If adoCmd("arg_e_code") <> "0" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        ss1.MaxRows = 0
        AEB2060C.Complete = True
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Private Sub CMD1_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    'Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Form_Ref
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    'Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
   ' Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
    Dim sTemp As String
    Dim iRow As Integer
    
    sQuery = "SELECT B.SLAB_EDT_SEQ," '--1
    sQuery = sQuery + "B.BLOCK_SEQ," '--2
    sQuery = sQuery + "B.SEQ ," '--3
    sQuery = sQuery + "B.ORD_NO," '--4
    sQuery = sQuery + "B.ORD_ITEM," '--5
    sQuery = sQuery + "B.ORD_CNT," '--6
    sQuery = sQuery + "B.OVER_FL," '--7
    sQuery = sQuery + "B.THK," '--8
    sQuery = sQuery + "B.WID," '--9
    sQuery = sQuery + "B.LEN," '--10
    sQuery = sQuery + "B.WGT," '--11
    sQuery = sQuery + "B.CNT," '--12
    sQuery = sQuery + "B.TRIM_FL," '--13
    sQuery = sQuery + "B.SMP_FL," '--14
    sQuery = sQuery + "B.SMP_LOC," '--15
    sQuery = sQuery + "B.SMP_LEN" '--16
    sQuery = sQuery + " FROM EP_PLATE_EDT B"
    sQuery = sQuery + " WHERE B.SLAB_EDT_SEQ = " & TXT_SLAB_NO.Value
      'sQuery = sQuery + " AND B.SLAB_EDT_SEQ = 1"
    sQuery = sQuery + "   "
    sQuery = sQuery + " Order by B.BLOCK_SEQ "
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

            If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
                'Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                For iRow = 1 To ss1.MaxRows
                
                    ss1.Row = iRow
                    ss1.Col = 2
                    sTemp = ss1.Text
                    ss1.Col = 3
                    
                    If sTemp = "00" And ss1.Text = "00" Then
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, RED)
                    ElseIf sTemp <> "00" And ss1.Text = "00" Then
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, BLUE)
                    Else
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow)
                    End If
                    
                Next iRow
            Else
                cmd_del.Enabled = False
            End If
            
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
            sMesg = sMesg + " Must input necessarily"
            Call Gp_MsgBoxDisplay(sMesg)
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

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub SSRibbon1_Click(Value As Integer)
    Unload Me
End Sub

Private Sub Form_Resize()
    
    If Me.Width - 1320 < 0 Or Me.Height - 1440 < 0 Then
        Me.Width = 1320
        Me.Height = 1440
        Exit Sub
    End If
    
    ss1.Width = Me.Width - 225
    ss1.Height = Me.Height - 1100
    Line1.X2 = ss1.Width
    CMD1.Left = Me.Width - 1320
    cmd_del.Left = CMD1.Left - cmd_del.Width - 10

End Sub

