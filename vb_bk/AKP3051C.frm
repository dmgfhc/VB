VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AKP3051C 
   Caption         =   "中厚板卷厂钢种分类维护_AKP3051C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_STLGRD_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10050
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.TextBox TXT_STLGRD 
      Height          =   285
      Left            =   8400
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ComboBox CBO_PROD_CD 
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
      ItemData        =   "AKP3051C.frx":0000
      Left            =   3315
      List            =   "AKP3051C.frx":000D
      TabIndex        =   3
      Text            =   "SL"
      Top             =   90
      Width           =   690
   End
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
      ItemData        =   "AKP3051C.frx":001D
      Left            =   1215
      List            =   "AKP3051C.frx":0024
      TabIndex        =   1
      Text            =   "C1"
      Top             =   90
      Width           =   690
   End
   Begin VB.ComboBox CBO_STDSPEC_GROUP 
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
      ItemData        =   "AKP3051C.frx":002C
      Left            =   5400
      List            =   "AKP3051C.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2775
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   90
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "工厂代码"
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
      Height          =   315
      Left            =   4260
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "钢种分类"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8610
      Left            =   90
      TabIndex        =   2
      Top             =   510
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15187
      _StockProps     =   64
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
      MaxCols         =   7
      MaxRows         =   20
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKP3051C.frx":0030
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   2190
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "产品分类"
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
Attribute VB_Name = "AKP3051C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      公辅材料是用实绩查询及修改界面
'-- Program ID        AGC2100C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2003.7.23
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
Public sDateTime As String          'Active Form Time Setting

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

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_PROD_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(CBO_STDSPEC_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKP3051C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AKP3051C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AKP3051C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 5, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub CBO_PROD_CD_Click()
    Call Gf_StlCboAdd(M_CN1, CBO_STDSPEC_GROUP, "C1", CBO_PROD_CD.Text)
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    With MDIMain.MenuTool
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

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

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 5)
    
    Call Gf_StlCboAdd(M_CN1, CBO_STDSPEC_GROUP, "C1", "SL")

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)

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

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        pControl(1).SetFocus
        CBO_PLT.Text = "C1"
        CBO_PROD_CD.Text = "SL"
        CBO_STDSPEC_GROUP.ListIndex = 0
    With MDIMain.MenuTool
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    With MDIMain.MenuTool
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    End If

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    With MDIMain.MenuTool
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub

Public Sub Form_Ins()

    If CBO_PROD_CD.Text <> "PP" And CBO_PROD_CD.Text <> "HC" And CBO_PROD_CD.Text <> "SL" Then
        Call Gp_MsgBoxDisplay("产品分类必须为 PP HC 或者 SL ")
        Exit Sub
    End If
    
    
    If CBO_STDSPEC_GROUP.Text = "" And CBO_PROD_CD.Text <> "SL" Then
        Call Gp_MsgBoxDisplay("请选择钢种分类")
        Exit Sub
    End If

    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.Col = 1
    ss1.Text = CBO_PLT.Text
    ss1.Col = 2
    ss1.Text = CBO_PROD_CD.Text
    ss1.Col = 3
    ss1.Text = CBO_STDSPEC_GROUP.Text
    ss1.Col = 6
    ss1.Text = sUserID

End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))
'
End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

    Call Gp_Sp_Del(Proc_Sc("SC"))

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
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub

    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub
Private Sub SS1_KeyUp(KeyCode As Integer, Shift As Integer)

    If ss1.Col = 4 Then
          
         If KeyCode = vbKeyF4 Then
         
            TXT_STLGRD.Text = ""
            TXT_STLGRD_NAME.Text = ""
         
            DD.sWitch = "MS"
            TXT_STLGRD.Text = ""
            DD.rControl.Add Item:=TXT_STLGRD
            DD.rControl.Add Item:=TXT_STLGRD_NAME
            
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
            ss1.Col = 4
            ss1.Text = TXT_STLGRD_NAME.Text
    
            Exit Sub
              
        End If
        
    End If

End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_MS_CommonNameFind
'   2.Name         : Common Code Name Find
'   3.Input  Value : Common Code , Code(TextBox) , CodeName(TextBox),CodeName1(TextBox)
'   4.Return Value :
'   5.Writer       : Chu Kyo Su
'   6.Create Date  : 2003. 10. 10
'   7.Modify Date  :
'   8.Comment      : Matser Type Common Code Name Find
'---------------------------------------------------------------------------------------
Public Sub Gp_MS_CodeNameFind(KeyCode As Integer, ByVal sCode As String, oCode As Object, Optional oCodeName As Object, Optional oCodeName1 As Object)
              
    Dim bType As Boolean
    Dim bCheck As Boolean
              
    If KeyCode = vbKeyF4 Then
                 
        DD.sWitch = "MS"
        DD.rControl.Add Item:=oCode
        DD.nameType = "2"
        
        If Not oCodeName Is Nothing Then DD.rControl.Add Item:=oCodeName
        If Not oCodeName1 Is Nothing Then DD.rControl.Add Item:=oCodeName1
        
        Select Case sCode
        
            Case "STDSPEC"              '标准号
                Call Gf_StdSPEC_DD(M_CN1, KeyCode)
            
            Case "CUST_CD"              '客户
                DD.nameType = "1"
                Call Gf_Customer_DD(M_CN1, KeyCode)
            
            Case "ENDUSE_CD"            '订单用途
                Call Gf_Usage_DD(M_CN1, KeyCode)
            
            Case "STLGRD"               '钢种
                Call Gf_Stlgrd_DD(M_CN1, KeyCode)
            
            Case "CUST_SPEC_NO"         '客户特殊要求编号
                Call Gf_Cust_STD_DD(M_CN1, KeyCode)
            
            Case "NISCO_QUALITY_NO"     '企标材质编号
                Call Gf_Nisco_STD_DD(M_CN1, KeyCode)
                
            Case "MLT_STD_NO"           '炼钢规程编号
                Call Gf_Melt_STD_DD(M_CN1, KeyCode)
                
            Case "MILL_STD_NO"          '轧钢规程编号
                Call Gf_Roll_STD_DD(M_CN1, KeyCode)
            
            Case "DEV_STD_CD"          '代表性交付条件标准
                Call Gf_STD_DELV_DD(M_CN1, KeyCode)
                            
            Case Else                   'Common Code
                DD.sKey = sCode
                Call Gf_Common_DD(M_CN1, vbKeyF4)
                bCheck = True
        
        End Select
        
    Else    'Max Length Input -> Code Name Find

        'If sCode = "" Then Exit Sub
        If KeyCode = 13 Or KeyCode = 20 Then Exit Sub
        
        If oCodeName Is Nothing Then Exit Sub
        
        Select Case sCode
        
            Case "STDSPEC"      '标准号
                bType = False
            Case Else           'Common Code
                bType = True
        End Select

        If bType = True And Len(Trim(oCode.Text)) = oCode.MaxLength Then
            If Left(oCodeName.Name, 3) = "lbl" And bCheck = True Then
                oCodeName.Caption = ""
                oCodeName.Caption = Gf_ComnNameFind(M_CN1, sCode, oCode.Text, "2")
            ElseIf bType = True Then
                oCodeName.Text = ""
                oCodeName.Text = Gf_ComnNameFind(M_CN1, sCode, oCode.Text, "2")
            End If
        ElseIf Len(Trim(oCode.Text)) = 0 Then
            If Left(oCodeName.Name, 3) = "lbl" Then
                oCodeName.Caption = ""
            Else
                oCodeName.Text = ""
            End If
        End If
        
    End If

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Plate_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant,sPRC String,
'                    {sFACT_CD,sPRC_LINE String, sADDNUM As Integer, ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : combo Add
'---------------------------------------------------------------------------------------
Public Function Gf_StlCboAdd(Conn As ADODB.Connection, Cbo As Variant, Optional sPlt As String = "C1", _
             Optional sProd_CD As String = "PP", Optional sADDNUM As Integer = 1000, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String

    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_StlCboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT STDSPEC_GROUP FROM (SELECT A.STDSPEC_GROUP "
    sQuery = sQuery + "               FROM GP_STDSPEC_GROUP A "
    sQuery = sQuery + "              WHERE A.PLT = '" + sPlt + "'"
    sQuery = sQuery + "                AND A.PROD_CD = '" + sProd_CD + "'"
    sQuery = sQuery + "           GROUP BY A.STDSPEC_GROUP  "
    sQuery = sQuery + "           ORDER BY A.STDSPEC_GROUP ASC) "
    sQuery = sQuery + "              WHERE ROWNUM <= " + CStr(sADDNUM)

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    Cbo.AddItem ""
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_StlCboAdd = True
    Else
        Gf_StlCboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_StlCboAdd = False

End Function
