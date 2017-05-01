VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ARZ0010C 
   Caption         =   "接口现状查询及实绩再发送_ARZ0010C"
   ClientHeight    =   8385
   ClientLeft      =   225
   ClientTop       =   2490
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   14760
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_status 
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
      Left            =   14700
      MaxLength       =   1
      TabIndex        =   6
      Tag             =   "进程状态"
      Top             =   150
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.ComboBox cbo_status 
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
      ItemData        =   "ARZ0010C.frx":0000
      Left            =   12045
      List            =   "ARZ0010C.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "进程状态"
      Top             =   120
      Width           =   2580
   End
   Begin VB.TextBox txt_if_cd 
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
      Left            =   1395
      TabIndex        =   0
      Tag             =   "库代码"
      Top             =   135
      Width           =   1050
   End
   Begin VB.TextBox txt_if_name 
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
      Left            =   2445
      MaxLength       =   40
      TabIndex        =   1
      Top             =   135
      Width           =   3435
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   150
      Top             =   135
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "接口种类"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8685
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   15165
      _Version        =   393216
      _ExtentX        =   26749
      _ExtentY        =   15319
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
      MaxCols         =   7
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ARZ0010C.frx":0067
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   6360
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "发送时间"
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
   Begin InDate.UDate udt_date_fr 
      Height          =   315
      Left            =   7605
      TabIndex        =   3
      Tag             =   "发送时间"
      Top             =   120
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
   Begin InDate.UDate udt_date_to 
      Height          =   315
      Left            =   9075
      TabIndex        =   4
      Tag             =   "发送时间"
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
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   10950
      Top             =   120
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "进程状态"
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
Attribute VB_Name = "ARZ0010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name
'-- Program Name
'-- Program ID        ARZ0010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHANG LIN
'-- Coder             ZHANG LIN
'-- Date              2008.03.3
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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_if_cd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_if_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_date_fr, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_date_to, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_status, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_status, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    
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
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cbo_status_Click()
    
    Select Case cbo_status.ListIndex
    
            Case 1
                txt_status.Text = "0"
            Case 2
                txt_status.Text = "1"
            Case 3
                txt_status.Text = "N"
            Case 4
                txt_status.Text = "S"
            Case 5
                txt_status.Text = "E"
            Case 6
                txt_status.Text = "D"
            Case 0
                txt_status.Text = ""
    End Select
    
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

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    
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

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        ss1.MaxCols = 7
        rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If Sp_Refer(M_CN1, ss1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
'        Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc").Item("Spread"))
    End If
            
End Sub

Public Sub Form_Pro()

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
    
'    If Gf_Sc_Authority(sAuthority, "U") Then
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'    End If
    
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

Private Sub txt_if_cd_DblClick()

   Call txt_if_cd_KeyUp(vbKeyF4, 0)
   
End Sub

Private Sub txt_if_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
       
        Load ARX0010C
        ARX0010C.txt_form_nm.Text = "ARZ0010C"
        ARX0010C.Show 1
    
    Else
    
        txt_if_name.Text = Gf_CodeFind(M_CN1, "SELECT IF_NAME FROM RP_IF_CD WHERE IF_ID = '" & txt_if_cd.Text & "'")
        
    End If
    
End Sub

Private Function Sp_Refer(Conn As ADODB.Connection, sPname As vaSpread, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim lStart As Long
    Dim lDataLen As Long
    Dim StartDate As Double
    Dim EndDate  As Double
    Dim sMsg  As String
    Dim sRSend   As String
    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords, ArrayRecords1 As Variant
    
    sMsg = Gf_Ms_NeceCheck(Mc1("nControl"))
    If sMsg <> "OK" Then
        sMsg = sMsg + "必须输入"
        Call Gp_MsgBoxDisplay(sMsg, "", "错误提示")
        Sp_Refer = False
        Exit Function
    End If
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Sp_Refer = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Refer = True
        
        .ReDraw = False
        .MaxCols = 7
        .MaxRows = 0: iCount = 0
        
        Screen.MousePointer = vbHourglass
        
        'RECEIVE/SEND
        sRSend = Gf_CodeFind(Conn, "SELECT IF_TYPE FROM RP_IF_CD WHERE IF_ID = '" & txt_if_cd.Text & "'")
        
        'Header Size SELECT
        sQuery = "          SELECT ITEM_NAME, START_PNT, DATA_LEN, DATA_TYPE, DATA_LEN2 FROM RP_IF_ITEM "
        sQuery = sQuery + "  WHERE IF_ID = '" & txt_if_cd.Text & "' ORDER BY ITEM_SEQ "
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Sp_Refer = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        
        .MaxCols = .MaxCols + UBound(ArrayRecords, 2) + 1
        
        For iRowCount = 7 To .MaxCols - 1
            
            .Row = 0
            .Col = iRowCount + 1
            
            If VarType(ArrayRecords(0, iRowCount - 7)) = vbNull Then
                .Text = ""
            Else
                .Text = Trim(ArrayRecords(0, iRowCount - 7))
            End If
            
            If Trim(ArrayRecords(3, iRowCount - 7)) = "N" Then
            
                .Col = iRowCount + 1: .Col2 = iRowCount + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = ArrayRecords(4, iRowCount - 7)
                .TypeNumberMax = 999999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroYes
                .BlockMode = False
            
            End If

        Next iRowCount
        
        .Col = 7: .Col2 = .MaxCols
        .Row = 1: .Row2 = -1
        .BlockMode = True
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
        
        'Start TIMESTAMP
        sQuery = "SELECT (TO_DATE('" & udt_date_fr.RawData & "000000', 'YYYY-MM-DD HH24:MI:SS') - TO_DATE('19700101080000',"
        sQuery = sQuery + " 'YYYY-MM-DD HH24:MI:SS')) * 86400000 + TO_NUMBER(TO_CHAR(SYSTIMESTAMP(3), 'FF')) FROM DUAL "
        
        StartDate = Gf_FloatFind(Conn, sQuery)
        
        'End TIMESTAMP
        sQuery = "SELECT (TO_DATE('" & udt_date_to.RawData & "235959', 'YYYY-MM-DD HH24:MI:SS') - TO_DATE('19700101080000',"
        sQuery = sQuery + "'YYYY-MM-DD HH24:MI:SS')) * 86400000 + TO_NUMBER(TO_CHAR(SYSTIMESTAMP(3), 'FF')) FROM DUAL "
        
        EndDate = Gf_FloatFind(Conn, sQuery)
        
        'Data Select
        If sRSend = "R" Then
                                        
            sQuery = "         SELECT  TIMESTAMP, SERIALNO, QUEUEID, HEADER, STATUS, PROCESSTIME, DESCRIPTION, SUBSTRB(DATA,11)  "
            sQuery = sQuery + "  FROM  NISCO.TBDIPDI "
            sQuery = sQuery + " WHERE  TIMESTAMP  BETWEEN " & StartDate & " AND " & EndDate
            sQuery = sQuery + "   AND  TRIM(SUBSTRB(DATA,1,10))    =      '" & txt_if_cd.Text & "' "
            sQuery = sQuery + "   AND  STATUS     LIKE   '" & txt_status.Text & "%' "
            sQuery = sQuery + " ORDER  BY  TIMESTAMP ASC, SERIALNO ASC "

        Else
        
            sQuery = "         SELECT  TIMESTAMP, SERIALNO, QUEUEID, HEADER, STATUS, PROCESSTIME, DESCRIPTION, SUBSTRB(DATA,11)  "
            sQuery = sQuery + "  FROM  NISCO.TBDIPDO "
            sQuery = sQuery + " WHERE  TIMESTAMP  BETWEEN " & StartDate & " AND " & EndDate
            sQuery = sQuery + "   AND  TRIM(SUBSTRB(DATA,1,10))    =      '" & txt_if_cd.Text & "' "
            sQuery = sQuery + "   AND  STATUS     LIKE   '" & txt_status.Text & "%' "
            sQuery = sQuery + " ORDER  BY  TIMESTAMP ASC, SERIALNO ASC "
            
        End If
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
            
            Sp_Refer = False
            .ReDraw = True
            .Refresh
            AdoRs.Close
            Set AdoRs = Nothing
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords1 = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
       .MaxRows = UBound(ArrayRecords1, 2) + 1

        For iRowCount = 0 To .MaxRows - 1

            .Row = iRowCount + 1
            
            For iColcount = 0 To 6
                .Col = iColcount + 1
                If VarType(ArrayRecords1(iColcount, iRowCount)) = vbNull Then
                    .Text = ""
                Else
                    .Text = RTrim(ArrayRecords1(iColcount, iRowCount))
                End If
            Next iColcount
            
            For iColcount = 7 To .MaxCols - 1

                .Col = iColcount + 1
                
                lStart = ArrayRecords(1, iColcount - 7)
                lDataLen = ArrayRecords(2, iColcount - 7)

                If VarType(ArrayRecords1(7, iRowCount)) = vbNull Then
                    .Text = ""
                Else
                    .Text = StrConv(MidB(StrConv(RTrim(ArrayRecords1(7, iRowCount)), vbFromUnicode), lStart, lDataLen), vbUnicode)
                End If

            Next iColcount

        Next iRowCount
        

        If Not lColumn Is Nothing Then

            'lControl Lock
            For iCount = 1 To lColumn.Count

                .Protect = True
                .Col = lColumn(iCount): .Col2 = lColumn(iCount)
                .Row = 1:               .Row2 = .MaxRows
                .BlockMode = True: .Lock = True
                .BlockMode = False

            Next iCount

        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Refer = False
    Call Gp_MsgBoxDisplay("Sp_Refer Error : " & sQuery)
    Screen.MousePointer = vbDefault

End Function
