VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form CEG3020C 
   Caption         =   "坯料使用计划的分析信息查询_CEG3020C"
   ClientHeight    =   9240
   ClientLeft      =   165
   ClientTop       =   1740
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   8685
      Left            =   30
      TabIndex        =   6
      Top             =   510
      Width           =   15225
      _Version        =   393216
      _ExtentX        =   26855
      _ExtentY        =   15319
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      SpreadDesigner  =   "CEG3020C.frx":0000
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   95
      Width           =   2175
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   1305
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   95
      Width           =   465
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   150
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt1 
      Height          =   315
      Left            =   8430
      TabIndex        =   2
      Top             =   90
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
      StartText.x     =   3
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
      BorderStyle     =   0
      FmtControl      =   1
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   7260
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "计划重量"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt2 
      Height          =   315
      Left            =   11130
      TabIndex        =   3
      Top             =   90
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
      StartText.x     =   3
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
      BorderStyle     =   0
      FmtControl      =   1
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   9960
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "申请重量"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt3 
      Height          =   315
      Left            =   13800
      TabIndex        =   4
      Top             =   90
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
      StartText.x     =   3
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
      BorderStyle     =   0
      FmtControl      =   1
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   12630
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "供给重量"
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
   Begin InDate.UDate udt_yymm 
      Height          =   315
      Left            =   5370
      TabIndex        =   5
      Tag             =   "指示日期"
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   134
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   4290
      Top             =   90
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "年月"
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
End
Attribute VB_Name = "CEG3020C"
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
'-- Program Name      SLAB USE PLAN
'-- Program ID        CEG3020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2007.10.24
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

    Dim iRow As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(udt_yymm, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_slab_wgt1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_slab_wgt2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_slab_wgt3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    ss1.Row = SpreadHeader + (ss1.ColHeaderRows - 2)
    ss1.RowHidden = True
    
End Sub

Public Sub Sp_Setting()

    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 2)) = 19
    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 1)) = 5
    ss1.MaxCols = 0

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
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Sp_Setting
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    txt_plt.Text = "C3"
    Call txt_plt_KeyUp(0, 0)
    
    Screen.MousePointer = vbDefault
    
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

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        txt_plt.Text = "C3"
        Call txt_plt_KeyUp(0, 0)
        ss1.MaxCols = 0
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Ref()

    Dim sQuery1 As String   'Header Display
    Dim sQuery2 As String   'Data Display
    Dim SMESG As String
    
    sdb_slab_wgt1.Value = 0
    sdb_slab_wgt2.Value = 0
    sdb_slab_wgt3.Value = 0
    
    'Header Display
    sQuery1 = "SELECT  DISTINCT  SLAB_THK "
    sQuery1 = sQuery1 + "  FROM  EP_REQ_SLAB_MPLAN "
    sQuery1 = sQuery1 + " WHERE  REQ_PLT     =  '" & txt_plt.Text & "'"
    sQuery1 = sQuery1 + "   AND  INCOM_YYMM  =  '" & udt_yymm.RawData & "'"
    sQuery1 = sQuery1 + " ORDER  BY SLAB_THK ASC "
    
    'Data Display
    sQuery2 = " {call CEG3020C.P_DATA ('" & txt_plt.Text & "','" & udt_yymm.RawData & "')} "

    SMESG = Gf_Ms_NeceCheck(nControl)
    If SMESG = "OK" Then
    
        SMESG = Gf_Ms_NeceCheck2(mControl)
        If SMESG = "OK" Then

            'Header Display
            Call Sp_Header_Refer1(ss1, sQuery1)      'Header Display
        
            'Data Display
            If Sp_Data_Refer1(ss1, sQuery2) Then     'Data Display
                ss1.OperationMode = OperationModeNormal
                Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
                ss1.Row = ss1.MaxRows
                ss1.Col = ss1.MaxCols - 2
                sdb_slab_wgt1.Value = IIf(ss1.Text = "", 0, ss1.Value)
                ss1.Col = ss1.MaxCols - 1
                sdb_slab_wgt2.Value = IIf(ss1.Text = "", 0, ss1.Value)
                ss1.Col = ss1.MaxCols
                sdb_slab_wgt3.Value = IIf(ss1.Text = "", 0, ss1.Value)
            End If
            
        Else
            Call Gp_MsgBoxDisplay(Trim(SMESG) + "长度不正确", "I")
        End If
    
    Else
        Call Gp_MsgBoxDisplay(Trim(SMESG) + "必须输入", "I")
    End If

End Sub

Public Sub Form_Pro()
    
End Sub

Public Sub Spread_Can()

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

Public Function Sp_Header_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Header_Refer1 = True
        
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer1 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 9
            For iCol = 0 To .MaxCols - 1 Step 9
            
                For iColCnt = 1 To 9
                
                    .Col = iCol + iColCnt
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 3)
                    
                    If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(0, iCnt))
                    End If
                    
                    .ColWidth(iCol + iColCnt) = 6
    
                    .Col = iCol + iColCnt: .Col2 = iCol + iColCnt
                    .Row = 1: .Row2 = -1
                    .BlockMode = True
                    .CellType = 13      'SS_CELL_TYPE_NUMBER
                    .TypeNumberDecPlaces = 0
                    .TypeNumberMax = 999999999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeNumberLeadingZero = TypeLeadingZeroYes

                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                    .BlockMode = False
                    
                    .Col = iCol + iColCnt
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    
                    Select Case iColCnt
                        Case 1, 2, 3
                            .Text = "板卷炼钢厂"
                        Case 4, 5, 6
                            .Text = "老炼钢厂"
                            .ColHidden = True
                        Case 7, 8, 9
                            .Text = "合计"
                            .ColHidden = True
                            Call Gp_Sp_ColHidden(ss1, .Col, True)
                    End Select
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    
                    Select Case iColCnt
                        Case 1, 4, 7
                            .Text = "计划"
                        Case 2, 5, 8
                            .Text = "申请"
                        Case 3, 6, 9
                            .Text = "供给"
                    End Select
                    
                Next iColCnt
                
                iCnt = iCnt + 1
                
            Next iCol
            
            '合计 Col
            For iColCnt = 1 To 9
                
                .MaxCols = .MaxCols + 1
                .Col = .MaxCols
                .Row = SpreadHeader + (.ColHeaderRows - 3)
                .Text = "合计"
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                    
                Select Case iColCnt
                    Case 1, 2, 3
                        .Text = "板卷炼钢厂"
                    Case 4, 5, 6
                        .Text = "老炼钢厂"
                        .ColHidden = True
                    Case 7, 8, 9
                        .Text = "合计"
                        .ColHidden = True
                End Select
                
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                Select Case iColCnt
                    Case 1, 4, 7
                        .Text = "计划"
                    Case 2, 5, 8
                        .Text = "申请"
                    Case 3, 6, 9
                        .Text = "供给"
                End Select
                    
                .ColWidth(.Col) = 6
                    
                .Col = .MaxCols: .Col2 = .MaxCols
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 999999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .TypeVAlign = TypeVAlignCenter
                .BlockMode = False
                
            Next iColCnt
            
        End If
        
        .BlockMode = True
        .Col = .MaxCols:  .Col2 = .MaxCols
        .Row = 1: .Row2 = -1
        .ForeColor = &HFF&  '&H00FF0000&
        .BlockMode = False
        
        For iColCnt = 9 To .MaxCols - 9 Step 9
            .BlockMode = True
            .Col = iColCnt - 2: .Col2 = iColCnt
            .Row = 1: .Row2 = -1
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iColCnt
        
        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 3)
        .Col = 1
        .Row2 = SpreadHeader + (.ColHeaderRows - 3)
        .Col2 = .MaxCols - 9
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False

        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 2)
        .Col = 1
        .Row2 = SpreadHeader + (.ColHeaderRows - 2)
        .Col2 = .MaxCols - 9
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False
        
        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 2)
        .Col = .MaxCols - 8
        .Row2 = SpreadHeader + (.ColHeaderRows - 1)
        .Col2 = .MaxCols
        .RowMerge = MergeAlways
        ''.ColMerge = MergeAlways
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
        
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    
    Dim iBas As Integer
    Dim iCot As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    Dim sStlgrd As String
    Dim sWid As String
    
    Dim ColSum(9) As Double
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer1 = True
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer1 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If iCnt = 0 Or sStlgrd <> Trim(ArrayRecords(0, iCnt)) Or sWid <> Trim(ArrayRecords(1, iCnt)) Then
                    sStlgrd = ArrayRecords(0, iCnt)
                    sWid = ArrayRecords(1, iCnt)
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = SpreadHeader + (.RowHeaderCols - 2)
                    .Text = Trim(ArrayRecords(0, iCnt))
                    .Col = SpreadHeader + (.RowHeaderCols - 1)
                    .Text = Trim(ArrayRecords(1, iCnt))
                End If
                
                For iCol = 1 To .MaxCols - 9 Step 9
                
                    .Col = iCol
                    .Row = SpreadHeader + (.ColHeaderRows - 3)
                    
                    If .Text = Trim(ArrayRecords(2, iCnt)) Then

                        .Row = .MaxRows
                        
                        For iColCnt = 1 To 9
                        
                            .Col = iCol + iColCnt - 1
                            If VarType(ArrayRecords(iColCnt + 2, iCnt)) = vbNull Then
                                .Text = ""
                            Else
                                If Trim(ArrayRecords(iColCnt + 2, iCnt)) = 0 Then
                                    .Text = ""
                                Else
                                    .Text = Trim(ArrayRecords(iColCnt + 2, iCnt))
                                End If
                            End If
                            
                        Next iColCnt
                            
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 0
        .Text = "合计(t)"
        
        'Column Sum
        For iCol = 1 To .MaxCols

            .Col = iCol

            If .Col <= 26 Then
                sCol_a = Chr(.Col + 64)
                .Formula = "sum(" + sCol_a + "1:" + sCol_a & .MaxRows - 1 & ")"
            Else
                iCot = Int(((.Col - 1) / 26))
                iBas = 26 * iCot
                sCol_a = Chr((.Col - iBas) + 64)
                sCol_b = Chr(iCot + 64)
                .Formula = "sum(" + sCol_b + sCol_a + "1:" + sCol_b + sCol_a & .MaxRows - 1 & ")"
            End If

        Next iCol

        'Row Sum
        For iRow = 1 To .MaxRows - 1

            .Row = iRow

            ColSum(1) = 0
            ColSum(2) = 0
            ColSum(3) = 0
            ColSum(4) = 0
            ColSum(5) = 0
            ColSum(6) = 0
            ColSum(7) = 0
            ColSum(8) = 0
            ColSum(9) = 0
            
            For iCol = 1 To .MaxCols - 9 Step 9

                For iColCnt = 1 To 9
                    .Col = iCol + iColCnt - 1
                    If .Text <> "" Then
                        ColSum(iColCnt) = ColSum(iColCnt) + .Value
                    End If
                Next iColCnt
                
            Next iCol

            For iColCnt = 9 To 1 Step -1
                .Col = .MaxCols - (iColCnt - 1)
                
                .Text = IIf(ColSum(10 - iColCnt) <> 0, ColSum(10 - iColCnt), "")
            Next iColCnt
            
        Next iRow
        
        Call Gp_Sp_EvenRowBackcolor(sPname, 1)
        
        .BlockMode = True
        .Row = .MaxRows:  .Row2 = .MaxRows
        .Col = 1: .Col2 = -1
        .ForeColor = &HFF&
        .BlockMode = False
        
        For iCol = 9 To .MaxCols - 9 Step 9
            .BlockMode = True
            .Col = iCol - 2: .Col2 = iCol
            .Row = .MaxRows: .Row2 = .MaxRows
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iCol
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function
