VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACF0080C 
   Caption         =   "�����ɱ����ݸ���_ACF0080C"
   ClientHeight    =   8955
   ClientLeft      =   285
   ClientTop       =   2325
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   14115
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_PLT 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "ACF0080C.frx":0000
      Left            =   6255
      List            =   "ACF0080C.frx":000D
      TabIndex        =   4
      Tag             =   "��������"
      Top             =   120
      Width           =   735
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8340
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   16020
      _ExtentX        =   28258
      _ExtentY        =   14711
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACF0080C.frx":001D
      Begin FPSpread.vaSpread ss1 
         Height          =   8340
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   16020
         _Version        =   393216
         _ExtentX        =   28257
         _ExtentY        =   14711
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
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
         MaxRows         =   37
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0080C.frx":004F
      End
   End
   Begin InDate.ULabel ULabel3 
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   0
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "��������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   1260
      TabIndex        =   1
      Tag             =   "��ʼ����"
      Top             =   120
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin InDate.UDate prod_date_to 
      Height          =   315
      Left            =   2790
      TabIndex        =   2
      Tag             =   "��������"
      Top             =   120
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   5040
      Top             =   120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "��������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   8280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��������"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   2670
      TabIndex        =   3
      Top             =   240
      Width           =   90
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
Attribute VB_Name = "ACF0080C"
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




Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_PLT = 1  '�ƻ�


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
     Call Gp_Ms_Collection(prod_date_to, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
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

'      For iRow = 5 To ss1.MaxCols
'        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
'      Next iRow
    

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
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
'    sc1.Item("Spread").Text = "��"


    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
    'Call Gp_Sp_ColHidden(ss2, 13, True)
    
    
    
'    sc2.Item("Spread").Col = 0
'    sc2.Item("Spread").Row = 0
'    sc2.Item("Spread").Text = "��"

    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    
    
    
'    Sc3.Item("Spread").Col = 0
'    Sc3.Item("Spread").Row = 0
'    Sc3.Item("Spread").Text = "��"
    
        
End Sub




Private Sub CBO_PLT_Change()

With ss1

If CBO_PLT.Text = "C2" Then

.ROW = 1
.Col = SS1_PLT:  .Text = "�����"

ElseIf CBO_PLT.Text = "C1" Then

.ROW = 1
.Col = SS1_PLT:  .Text = "�к��"

ElseIf CBO_PLT.Text = "C3" Then

.ROW = 1
.Col = SS1_PLT:  .Text = "�а�"

End If

End With

End Sub

Private Sub CBO_PLT_Click()
With ss1

If CBO_PLT.Text = "C2" Then

.ROW = 1
.Col = SS1_PLT:  .Text = "�����"

ElseIf CBO_PLT.Text = "C1" Then

.ROW = 1
.Col = SS1_PLT:  .Text = "�к��"

ElseIf CBO_PLT.Text = "C3" Then

.ROW = 1
.Col = SS1_PLT:  .Text = "�а�"

End If

End With
End Sub

Private Sub Cmd_Edit_Click()
'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
          
    If Trim(prod_date_to) = "" Then
        Call Gp_MsgBoxDisplay(prod_date_to.Tag + "��������")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACF0080P ('" + Trim(Format(prod_date_from.Text, "YYYYMMDD")) + "','" + Trim(Format(prod_date_to.Text, "YYYYMMDD")) + "',?)}"

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
        
        Call Gp_MsgBoxDisplay("���³ɹ�..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("����ʧ�ܣ���")
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
Public Sub Form_Pro()

Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCount As Integer
    Dim var1 As String '����
    Dim var2 As String '������
    
    iCount = 0
    
          
    If Trim(prod_date_to.Text) = "____-__-__" Then
        Call Gp_MsgBoxDisplay(prod_date_to.Tag + "��������")
        Exit Sub
    End If
    
    If Trim(CBO_PLT.Text) = "" Then
        Call Gp_MsgBoxDisplay(CBO_PLT.Tag + "��������")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    For iRow = 5 To 7
    
    If iRow = 5 Then
    
    var1 = "CUT_WGT"
    
    ElseIf iRow = 6 Then
    
    var1 = "IRON_SCALE_WGT"
    
    ElseIf iRow = 7 Then
    
    var1 = "SCRAP_WGT"
    
    End If
    
    With ss1
    
    .ROW = iRow
    
    .Col = 0
    
    If .Text = "Update" Then
    
    .Col = 4: var2 = CStr(.Text)
    
    sQuery = "{call ACF0081P ('" + Trim(Format(prod_date_to.Text, "YYYYMMDD")) + "','" + Trim(CBO_PLT.Text) + "','" + var1 + "','" + var2 + "',?)}"

    'Ado Setting
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
            
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    iCount = iCount + 1
    
    .Col = 0: .Text = iRow
    
    End If
    
    End With
    
    Next iRow
    
    
    'Process Error Check

        'strRet_Result_ErrMsg = "��������ȷ�ı�������" 'adoCmd("arg_e_msg")
'        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
'        Call Gp_MsgBoxDisplay(sErrMessg)
    'Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ���ɹ�������" & iCount & "����¼"
        Call Form_Ref
        Exit Sub
    
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("����ʧ�ܣ���")
    
    

End Sub



Private Sub Form_Load()

Dim iRow As Integer
Dim i As Integer

i = 5
 
 Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
        
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    

    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

      ' Cmd_Edit.Enabled = True


    'txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
    
    Screen.MousePointer = vbDefault
    
    CBO_PLT.Text = "C2"
    
        With ss1

        For iRow = i To 7 Step 1
            
        .ROW = 5: .Col = 4
        .BlockMode = True
        .Lock = False
        '.BackColor = &HFFFFFF
        .BlockMode = False

        Next iRow


        End With
        
        prod_date_from.Text = Date - 1
        prod_date_to.Text = Date - 1
        

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)


    
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
    
    
    Set Mc1 = Nothing
    Set sc1 = Nothing

    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub



Public Sub Form_Cls()
    
Call Form_Ref
        
        
    
End Sub

Public Sub Form_Ref()

    If Trim(prod_date_from.RawData) = "" Then
        Call Gp_MsgBoxDisplay(prod_date_from.Tag + "��������")
        Exit Sub
    End If
    
    If Trim(prod_date_to.RawData) = "" Then
        Call Gp_MsgBoxDisplay(prod_date_to.Tag + "��������")
        Exit Sub
    End If
    
    If Trim(CBO_PLT.Text) = "" Then
        Call Gp_MsgBoxDisplay(CBO_PLT.Tag + "��������")
        Exit Sub
    End If
    
    ss1.ReDraw = False
    
    Call Zero_Cls
    
'    Call Form_Cls
    Screen.MousePointer = vbHourglass
        
    Call Ss1_Data_Refer
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    ss1.ReDraw = True
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    
    Screen.MousePointer = vbDefault
    
                        
End Sub

Public Sub Ss1_Data_Refer()

On Error GoTo Ss1_Display_Error


    
    Dim PLATE_WGT  As Double
    Dim slab_input_wgt   As Double
    Dim recyle_wgt   As Double
    Dim rolled_unplan_wgt As Double
    Dim rolled_wgt As Double
    Dim flat_unplan_wgt As Double
    Dim flat_wgt As Double
    Dim hcr_wgt As Double
    Dim BED_WGT As Double
    
    
    
    

    Dim AdoRs As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
  

    sQuery = "SELECT SUM(PLATE_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(BED_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(SLAB_INPUT_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(CUT_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(IRON_SCALE_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(SCRAP_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(GRIND_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(ROLL_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(ROLLED_UNPLAN_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(ROLLED_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(FLAT_UNPLAN_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(FLAT_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(SLAB_SCRAP_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(HCR_WGT)" & vbCrLf
    sQuery = sQuery & ",SUM(BLIND_TIME)" & vbCrLf
    sQuery = sQuery & ",SUM(ELECTRICAL_TIME)" & vbCrLf
    sQuery = sQuery & ",SUM(MACHINE_TIME)" & vbCrLf
    sQuery = sQuery & ",SUM(PLAN_TIME)" & vbCrLf
    sQuery = sQuery & ",SUM(OUT_TIME)" & vbCrLf
    sQuery = sQuery & ",SUM(OPERATION_TIME)" & vbCrLf
    sQuery = sQuery & ",SUM(OTHERS_TIME)" & vbCrLf
    sQuery = sQuery & "   FROM  GP_RPT_PRODUCTION_COST                  " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  BETWEEN  '" & prod_date_from.RawData & "' AND '" & prod_date_to.RawData & "'" & vbCrLf
    sQuery = sQuery & "    AND  PLT        =  '" & CBO_PLT.Text & "'" & vbCrLf
    
    

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Do Until AdoRs.EOF
       
        With ss1

            .Col = 4:   .ROW = 1:    .Text = Val(AdoRs.Fields(0) & ""): PLATE_WGT = .Value '�Ĳ���
                        .ROW = 2:    .Text = Val(AdoRs.Fields(1) & ""): BED_WGT = .Value '�����
                        .ROW = 3:    .Text = Val(AdoRs.Fields(2) & ""): slab_input_wgt = .Value '����Ͷ����
                        
                          If slab_input_wgt <> 0 Then
                        
                         .ROW = 4:    .Text = PLATE_WGT * 100 / slab_input_wgt '�ɲ���
                        
                          End If
                        
                        .ROW = 5:    .Text = Val(AdoRs.Fields(3) & "") '��˿������
                        .ROW = 6:    .Text = Val(AdoRs.Fields(4) & "") '������Ƥ������
                        .ROW = 7:    .Text = Val(AdoRs.Fields(5) & "") '�ϴ�Ʒ��������
                        
                        .ROW = 8:    .Text = Val(AdoRs.Fields(3) & "") + Val(AdoRs.Fields(4) & "") + Val(AdoRs.Fields(5) & ""): recyle_wgt = .Value 'С��
                        
                        
                        If slab_input_wgt <> 0 And PLATE_WGT <> 0 Then
                        
                        .ROW = 9:    .Text = (1 / (1 / (PLATE_WGT / slab_input_wgt) - recyle_wgt / PLATE_WGT)) * 100 '����ƽ��
                        
                        End If
                        
                         If PLATE_WGT <> 0 Then
                        
                        .ROW = 11:    .Text = Val(AdoRs.Fields(6) & "") / PLATE_WGT '����
                        
                        End If
                        
                        
                        .ROW = 10:   .Text = Val(AdoRs.Fields(6) & "") '����ĥ����
                        
                        
                        
                        .ROW = 12:    .Text = Val(AdoRs.Fields(8) & ""): rolled_unplan_wgt = .Value '�����Ǽƻ���
                        .ROW = 13:    .Text = Val(AdoRs.Fields(9) & ""): rolled_wgt = .Value '��������
                        
                        If rolled_wgt <> 0 Then
                        
                        .ROW = 14:    .Text = rolled_unplan_wgt * 100 / rolled_wgt '�����Ǽƻ���
                        
                         End If
                        
                        .ROW = 15:    .Text = Val(AdoRs.Fields(10) & ""): flat_unplan_wgt = .Value 'ƽ���Ǽƻ���
                        .ROW = 16:    .Text = Val(AdoRs.Fields(11) & ""): flat_wgt = .Value 'ƽ������
                        
                         If flat_wgt <> 0 Then
                        
                        .ROW = 17:    .Text = flat_unplan_wgt * 100 / flat_wgt 'ƽ���Ǽƻ���
                        
                        End If
                        
                        If BED_WGT <> 0 Then
                        
                        .ROW = 18:   .Text = (rolled_unplan_wgt + flat_unplan_wgt) * 100 / BED_WGT ' �ۺϷǼƻ���
                        
                        End If
                        
                        .ROW = 19:   .Text = Val(AdoRs.Fields(12) & "") '�����з����������з�����
                        
                        .ROW = 20:   .Text = Val(AdoRs.Fields(13) & ""): hcr_wgt = .Value '��װ������
                        
                        If slab_input_wgt <> 0 Then
                        
                        .ROW = 21:    .Text = hcr_wgt * 100 / slab_input_wgt '����
                        
                        End If
                        
                        .ROW = 22:    .Text = Val(AdoRs.Fields(14) & "") '�°�
                        .ROW = 23:    .Text = Val(AdoRs.Fields(15) & "") '����
                        .ROW = 24:    .Text = Val(AdoRs.Fields(16) & "") '��е
                        .ROW = 25:    .Text = Val(AdoRs.Fields(17) & "") '�ƻ�
                        .ROW = 26:    .Text = Val(AdoRs.Fields(18) & "") '�ⲿ
                        .ROW = 27:    .Text = Val(AdoRs.Fields(19) & "") '����
                        .ROW = 28:    .Text = Val(AdoRs.Fields(20) & "") '����
        
'                  -----------------------------------------20160808
                        
                        


        End With
    
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
    sQuery = "SELECT DZ_WGT" & vbCrLf
    sQuery = sQuery & ",QA_WGT" & vbCrLf
    sQuery = sQuery & ",XAC_WGT" & vbCrLf
    sQuery = sQuery & ",XAA_WGT" & vbCrLf
    sQuery = sQuery & ",HOT_WGT" & vbCrLf
    sQuery = sQuery & ",COLD_WGT" & vbCrLf
    sQuery = sQuery & ",POLISHED_WGT" & vbCrLf
    sQuery = sQuery & ",TRIM_WGT" & vbCrLf
    sQuery = sQuery & ",DEFECT_WGT " & vbCrLf
    sQuery = sQuery & "   FROM  GP_RPT_PRODUCTION_COST                  " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  =  '" & prod_date_to.RawData & "'" & vbCrLf
    sQuery = sQuery & "    AND  PLT        =  '" & CBO_PLT.Text & "'" & vbCrLf
    
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Do Until AdoRs.EOF
       
        With ss1

            .Col = 4:   .ROW = 29:   .Text = Val(AdoRs.Fields(0) & "") 'DZB/DZE
                        .ROW = 30:   .Text = Val(AdoRs.Fields(1) & "") 'QAB/QAE
                        .ROW = 31:   .Text = Val(AdoRs.Fields(2) & "") 'XAC
                        .ROW = 32:    .Text = Val(AdoRs.Fields(3) & "") 'XAA
                        .ROW = 33:    .Text = Val(AdoRs.Fields(4) & "") '���ȴ�����
                        .ROW = 34:    .Text = Val(AdoRs.Fields(5) & "") '����
                        .ROW = 35:    .Text = Val(AdoRs.Fields(6) & "") '����ĥ
                        .ROW = 36:    .Text = Val(AdoRs.Fields(7) & "") '���и�
                        .ROW = 37:    .Text = Val(AdoRs.Fields(8) & "") '�а������
        
        
'                  -----------------------------------------20160808
                        
                        


        End With
    
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
    With ss1
    
    .Col = 0
    
    .ROW = 5: .Text = 5
    .ROW = 6: .Text = 6
    .ROW = 7: .Text = 7
    
    End With
    
    
    Exit Sub

Ss1_Display_Error:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Ss1_Display_Error : " & Error)
    
End Sub

Public Sub Zero_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.ROW = iRow
            ss1.Col = 4
                ss1.Text = ""
    Next iRow

End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Form_Exc()


        Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub


Private Sub ss1_Change(ByVal Col As Long, ByVal ROW As Long)

If Col = 4 Then
    If (ROW = 5 Or ROW = 6 Or ROW = 7) Then

        ss1.ROW = ss1.ActiveRow
        ss1.Col = 0
        ss1.Text = "Update"
    End If
    
    End If

End Sub



Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

If Col = 4 Then
    If (ROW = 5 Or ROW = 6 Or ROW = 7) Then

        ss1.ROW = ss1.ActiveRow
        ss1.Col = 0
        ss1.Text = "Update"
    End If
    
    End If

End Sub