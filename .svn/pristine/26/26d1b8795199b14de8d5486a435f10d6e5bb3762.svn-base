VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFO3010C 
   Caption         =   "废钢代码下达界面(仅供转炉二级)_AFO3010C"
   ClientHeight    =   9225
   ClientLeft      =   315
   ClientTop       =   2280
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_con_id 
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
      ItemData        =   "AFO3010C.frx":0000
      Left            =   6750
      List            =   "AFO3010C.frx":0002
      TabIndex        =   4
      Tag             =   "炉座号"
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cbo_to 
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
      ItemData        =   "AFO3010C.frx":0004
      Left            =   4095
      List            =   "AFO3010C.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cbo_from 
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
      ItemData        =   "AFO3010C.frx":0008
      Left            =   1440
      List            =   "AFO3010C.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8670
      Left            =   75
      TabIndex        =   3
      Top             =   510
      Width           =   15090
      _Version        =   393216
      _ExtentX        =   26617
      _ExtentY        =   15293
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AFO3010C.frx":000C
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   2850
      Top             =   120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "至序号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   225
      Top             =   120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "始序号"
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
   Begin Threed.SSCommand cmd_send 
      Height          =   390
      Left            =   13260
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   688
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&下达到转炉二级"
   End
   Begin InDate.ULabel ULabel50 
      Height          =   315
      Left            =   5505
      Top             =   120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "炉座号"
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
End
Attribute VB_Name = "AFO3010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      Sending Scrap Code to BOF L2
'-- Program ID        AFO3010C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2004.8.7
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

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread Necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
        
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cbo_from_Click()
    Dim iRow, iCol As Integer
    Dim sColor As String
    ss1.SetFocus
    
    If cbo_to.Text <> "" And cbo_from.Text > cbo_to.Text Then
       MsgBox "始序号必须小于至序号！", vbCritical, "系统提示信息"
       cbo_from.SetFocus
       Exit Sub
    Else
       If cbo_to.Text <> "" Then
            With ss1
                 For iRow = 1 To CInt(cbo_to.Text) - 1
                     .Col = 2
                     .Row = iRow
                     sColor = .BackColor
                     
                     .Col = 1
                     .BackColor = sColor
                     
                     If .Row <= CInt(cbo_to.Text) And .Row >= CInt(cbo_from.Text) Then
                        .BackColor = &HFF
                     End If
                 Next iRow
            End With
       Else
            With ss1
                 For iRow = 1 To .MaxRows
                     .Row = iRow
                     .Col = 2
                     sColor = .BackColor
                     
                     .Col = 1
                     .BackColor = sColor
                     
                     If .Text = cbo_from.Text Then
                        .BackColor = &HFF
                     End If
                 Next iRow
            End With
       End If
    End If

End Sub

Private Sub cbo_to_Click()
    Dim iRow, iCol As Integer
    Dim sColor As String
    
    If cbo_from = "" Then
       MsgBox "请先选择始序号！", vbCritical, "系统提示信息"
       cbo_from.SetFocus
       Exit Sub
    Else
       If cbo_to.Text < cbo_from.Text Then
          MsgBox "至序号必须大于始序号！", vbCritical, "系统提示信息"
          cbo_to.SetFocus
          Exit Sub
       Else
            With ss1
                 For iRow = CInt(cbo_from.Text) + 1 To .MaxRows
                     .Col = 2
                     .Row = iRow
                     sColor = .BackColor
                     
                     .Col = 1
                     .BackColor = sColor
                     
                     If .Row <= CInt(cbo_to.Text) Then
                        .BackColor = &HFF
                     End If
                 Next iRow
                 
            End With
       End If
    
    End If

End Sub

Private Sub cmd_send_Click()
    Dim iRow As Integer
    Dim sColor As String
    
    If cbo_from.Text = "" Or cbo_to.Text = "" Then
       MsgBox "请选择始序号和至序号后再下达！", vbCritical, "系统提示信息"
       Exit Sub
    Else
       
       If Gf_MessConfirm("确定要将从序号 <" + cbo_from.Text + "> 到序号 <" + cbo_to.Text + "> 的废钢代码下达到转炉二级吗？", "", "系统提示信息") Then
          Call Gp_Scrap_Send
          With ss1
               
               For iRow = 1 To .MaxRows
                   .Row = iRow
                   .Col = 2
                   sColor = .BackColor
                   .Col = 1
                   .BackColor = sColor
                   
               Next iRow
          End With
       Else
          Exit Sub
       End If
    End If
End Sub

Public Sub Gp_Scrap_Send()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    If Trim(cbo_con_id.Text) = "" Then cbo_con_id.Text = "1"
    
    sQuery = "{call AFO3010P ('" + cbo_from + "', '" + cbo_to + "')}"
    
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
        
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        Call MsgBox("您选择的废钢代码已成功下达到转炉二级！", vbInformation, "系统提示信息")
        'Call Form_Ref
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
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
        
    cbo_con_id.Clear
    cbo_con_id.AddItem "1"
    cbo_con_id.AddItem "2"
    cbo_con_id.AddItem "3"
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)

    cbo_con_id = "1"
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
        
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set iSumCol = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        cmd_send.Enabled = False
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim iRow As Integer
    Dim sTemp, sColor As String
    
    sQuery = "select ROWNUM,cd, cd_short_name from zp_cd where cd_mana_no = 'F0017'"
    sQuery = sQuery + " Order by ROWNUM"
    
    If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       MDIMain.StatusBar1.Panels(1) = "提示信息：查询成功"
       cmd_send.Enabled = True
       ss1.OperationMode = OperationModeNormal
       
       With ss1

         .Col = 1
    
         cbo_from.Clear
         cbo_from.AddItem ""
         cbo_to.Clear
         cbo_to.AddItem ""
    
         For iRow = 1 To .MaxRows
             .Row = iRow
             
            If sTemp <> .Text Then
        
               cbo_from.AddItem .Text
               cbo_to.AddItem .Text
                
               sTemp = .Text
            
            End If
                
            If iRow = 1 Then sTemp = .Text
            
         Next iRow
         
       End With

    End If
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
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

'Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
'
'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub

'Private Sub ss1_LostFocus()
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub


