VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGE2040C 
   Caption         =   "钢板库库图现状查询_CGE2040C"
   ClientHeight    =   9090
   ClientLeft      =   360
   ClientTop       =   1725
   ClientWidth     =   14205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14205
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_CUR_INV 
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
      ItemData        =   "CGE2040C.frx":0000
      Left            =   1380
      List            =   "CGE2040C.frx":000D
      TabIndex        =   7
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1680
   End
   Begin VB.ComboBox CBO_ZONE_TYPE 
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
      ItemData        =   "CGE2040C.frx":001D
      Left            =   6870
      List            =   "CGE2040C.frx":001F
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
   Begin VB.ComboBox CBO_YARD_TYPE 
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
      ItemData        =   "CGE2040C.frx":0021
      Left            =   6060
      List            =   "CGE2040C.frx":0023
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin VB.ComboBox CBO_YARD_ROW 
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
      ItemData        =   "CGE2040C.frx":0025
      Left            =   7620
      List            =   "CGE2040C.frx":0027
      TabIndex        =   1
      Top             =   9840
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ComboBox CBO_YARD_COLUMN 
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
      ItemData        =   "CGE2040C.frx":0029
      Left            =   8430
      List            =   "CGE2040C.frx":002B
      TabIndex        =   0
      Top             =   9840
      Visible         =   0   'False
      Width           =   795
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   12750
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "钢板总数"
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
      Height          =   8625
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   15270
      _Version        =   393216
      _ExtentX        =   26935
      _ExtentY        =   15214
      _StockProps     =   64
      DisplayColHeaders=   0   'False
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
      MaxCols         =   0
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGE2040C.frx":002D
      UserResize      =   1
   End
   Begin CSTextLibCtl.sidbEdit sdb_plate_cnt 
      Height          =   315
      Left            =   14010
      TabIndex        =   6
      Top             =   120
      Width           =   1290
      _Version        =   262145
      _ExtentX        =   2275
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
      RawData         =   ""
      Text            =   " 0"
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
      NumDecDigits    =   0
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   70
      Top             =   120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "当前库"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   4800
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "跨 / 区"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   6360
      Top             =   9840
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "行 / 列"
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
Attribute VB_Name = "CGE2040C"
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
'-- Program Name      钢板库库图现状查询
'-- Program ID        AGE2040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
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

Public Active_CForm As String       'Form Active
Public sOrd_no As String
Public Form_Wid As Double
Public Form_Len As Double

Dim rib_prev As Integer
Dim opt_prev As Integer

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

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
    
    Call Gp_Ms_Collection(CBO_CUR_INV, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_YARD_TYPE, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_ZONE_TYPE, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    Call Gp_Ms_Collection(sdb_plate_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

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
    
    rib_prev = 0
    opt_prev = 0
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Activate()
    
    If Active_CForm <> "" Then
        
        rib_prev = 0
        opt_prev = 0
        
        Call Form_Ref
        Active_CForm = ""
        
    End If
    
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
    
    Dim iInt As Integer
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    For iInt = 1 To 20
        cbo_YARD_TYPE.AddItem iInt
        CBO_ZONE_TYPE.AddItem Chr(64 + iInt)
'        CBO_YARD_ROW.AddItem Format(iInt, "00")
'        CBO_YARD_COLUMN.AddItem Format(iInt, "00")
    Next iInt
    
    CBO_CUR_INV.Text = "ZB"
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
    Dim sCur_Inv As String
    Dim sArea As String
    Dim sAddr As String
    Dim sYard_Row As String
    Dim sYard_Column As String
    Dim sPlate_no As String

    sMesg = Gf_Ms_NeceCheck(nControl)
    
    ss1.MaxRows = 0
    
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
        
            sCur_Inv = CBO_CUR_INV.Text
            sPlate_no = "00000000000000"
            
            sAddr = "P" & cbo_YARD_TYPE.Text & CBO_ZONE_TYPE.Text
            
            'MaxRows Setting
            sQuery = "         SELECT MAX(YARD_ROW) FROM FP_STDYARD "
            sQuery = sQuery + " WHERE LOCATION LIKE  '" + sAddr + "%' AND YARD_KND = '" + sCur_Inv + "' "
            ss1.MaxRows = Val(Gf_CodeFind(M_CN1, sQuery)) * 2


            'MaxCols Setting
            sQuery = "         SELECT MAX(YARD_COLUMN) FROM FP_STDYARD "
            sQuery = sQuery + " WHERE LOCATION LIKE  '" + sAddr + "%' AND YARD_KND = '" + sCur_Inv + "' "
            ss1.MaxCols = Val(Gf_CodeFind(M_CN1, sQuery))

            'Address Setting
            sQuery = "         SELECT LOCATION FROM FP_STDYARD "
            sQuery = sQuery + " WHERE LOCATION LIKE  '" + sAddr + "%' AND YARD_KND = '" + sCur_Inv + "' "
            sQuery = sQuery + " ORDER BY LOCATION "

            If Addr_Display(M_CN1, Proc_Sc("Sc"), sQuery, False) = False Then
                Exit Sub
            End If

            sQuery = "         SELECT A.YARD_ADDR, '('||A.CNT||','||A.BED_SEQ||')', (SELECT LEN || ' X ' || WID FROM GP_PLATE WHERE PLATE_NO = B.PLATE_NO ) PLATE_SIZE  "
            sQuery = sQuery + "  FROM GP_PLATEYARD B, "
            sQuery = sQuery + "        (SELECT MAX(Y.YARD_KND) YARD_KND,YARD_ADDR YARD_ADDR,COUNT(*) CNT,MAX(BED_SEQ) BED_SEQ "
            sQuery = sQuery + "           FROM GP_PLATEYARD Y  "
            sQuery = sQuery + "          WHERE Y.YARD_KND  = '" + sCur_Inv + "' "
            sQuery = sQuery + "            AND Y.YARD_ADDR  LIKE '" + sAddr + "%' "
            sQuery = sQuery + "            AND Y.PLATE_NO IS NOT NULL "
            sQuery = sQuery + "       GROUP BY Y.YARD_ADDR) A "
            sQuery = sQuery + " WHERE B.YARD_KND = '" + sCur_Inv + "' "
            sQuery = sQuery + " AND B.YARD_ADDR  LIKE '" + sAddr + "%' "
            sQuery = sQuery + " AND B.YARD_ADDR = A.YARD_ADDR  "
            sQuery = sQuery + " AND B.BED_SEQ   = A.BED_SEQ  "
            sQuery = sQuery + " AND B.YARD_KND  = A.YARD_KND "
'            sQuery = sQuery + " AND SUBSTR(B.YARD_ADDR,4,2)  > '00' "
'            sQuery = sQuery + " AND SUBSTR(B.YARD_ADDR,6,2)  > '00' "
            sQuery = sQuery + " ORDER BY A.YARD_ADDR"
            
'            sQuery = "          SELECT C.ADDR, '('||C.CNT||')', D.LEN, D.WID, D.ORD_NO  "
'            sQuery = sQuery + "   FROM (SELECT YARD_ADDR ADDR, COUNT(*) CNT  "
'            sQuery = sQuery + "           FROM GP_PLATEYARD  "
'            sQuery = sQuery + "          WHERE trim(PLATE_NO) is not null "
'            sQuery = sQuery + "            AND SUBSTR(YARD_ADDR,2,2) = '" + sAddr + sArea + "'  "
'            sQuery = sQuery + "          GROUP BY YARD_ADDR) C,  "
'            sQuery = sQuery + "        (SELECT A.LEN LEN, A.WID WID, B.YARD_ADDR ADDR, A.ORD_NO ORD_NO "
'            sQuery = sQuery + "           FROM GP_PLATE A, GP_PLATEYARD B  "
'            sQuery = sQuery + "          WHERE A.PLATE_NO IN  "
'            sQuery = sQuery + "                         (SELECT PLATE_NO  FROM GP_PLATEYARD  "
'            sQuery = sQuery + "                           WHERE YARD_ADDR||BED_SEQ IN  "
'            sQuery = sQuery + "                                                   (SELECT YARD_ADDR||MAX(BED_SEQ)  "
'            sQuery = sQuery + "                                                      FROM GP_PLATEYARD  "
'            sQuery = sQuery + "                                                     WHERE trim(PLATE_NO) is not null  "
'            sQuery = sQuery + "                                                       AND SUBSTR(YARD_ADDR,2,2) = '" + sAddr + sArea + "'  "
'            sQuery = sQuery + "                                                     GROUP BY YARD_ADDR ))  "
'            sQuery = sQuery + "            AND A.PLATE_NO = B.PLATE_NO) D  "
'            sQuery = sQuery + " WHERE C.ADDR = D.ADDR  "
'            sQuery = sQuery + " ORDER BY C.ADDR  "

'            sQuery = "                  SELECT  B.YARD_ADDR ADDR, '('||B.BED_SEQ||')', A.LEN LEN, A.WID WID, A.ORD_NO ORD_NO  "
'            sQuery = sQuery + "           FROM GP_PLATE A, GP_PLATEYARD B   "
'            sQuery = sQuery + "          WHERE A.PLATE_NO IN   "
'            sQuery = sQuery + "                         (SELECT PLATE_NO  FROM GP_PLATEYARD  "
'            sQuery = sQuery + "                           WHERE YARD_ADDR||BED_SEQ IN  "
'            sQuery = sQuery + "                                                   (SELECT YARD_ADDR||MAX(BED_SEQ)  "
'            sQuery = sQuery + "                                                      FROM GP_PLATEYARD  "
'            sQuery = sQuery + "                                                     WHERE trim(PLATE_NO) is not null"
'            sQuery = sQuery + "                                                       AND SUBSTR(YARD_ADDR,2,2) = '" + sAddr + sArea + "'  "
'            sQuery = sQuery + "                                                     GROUP BY YARD_ADDR ))  "
'            sQuery = sQuery + "            AND A.PLATE_NO = B.PLATE_NO   "
'            sQuery = sQuery + "       ORDER BY B.YARD_ADDR  "


            If Only_Display(M_CN1, Proc_Sc("Sc"), sQuery, False) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                sdb_plate_cnt.Value = Gf_FloatFind(M_CN1, _
                                     "SELECT COUNT(*) FROM GP_PLATEYARD WHERE YARD_KND = '" + sCur_Inv + "' " + "AND YARD_ADDR LIKE '" _
                                      + sAddr + "%' AND PLATE_NO IS NOT NULL ")
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



Public Sub Sp_Setting(ByVal sPname As Variant)

    With sPname
    
        .RowHeight(-1) = 18
        .ColWidth(-1) = 14
        
        .BackColorStyle = BackColorStyleUnderGrid
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        '.ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .UserResize = UserResizeNone
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .CellType = SS_CELL_TYPE_BUTTON
        .TypeButtonShadowSize = 3
        '.TypeButtonColor = vbWhite
        .BlockMode = False
        
        .MaxRows = 0
        
    End With
    
End Sub

Public Function Only_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn
    
    Dim lCount As Long
    Dim lCol As Long
    Dim lRow As Long
    
    Dim sArea As String
    Dim sAddr As String
    Dim sTemp As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Only_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
        
    With Sc.Item("Spread")

        Only_Display = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("没有相关的数据", "I")
                
            Only_Display = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = 0
            
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 2) + 1 > 0 Then 'Row Count
            
            For lCount = 0 To UBound(ArrayRecords, 2)
                
                If Mid(Trim(ArrayRecords(0, lCount)), 4, 2) > "00" And Mid(Trim(ArrayRecords(0, lCount)), 6, 2) > "00" Then
                    If sTemp <> Mid(Trim(ArrayRecords(0, lCount)), 4, 2) Then
                        sTemp = Mid(Trim(ArrayRecords(0, lCount)), 4, 2)
                    End If
                    
                    .Row = Val(sTemp) * 2 - 1
                    .Col = Val(Mid(Trim(ArrayRecords(0, lCount)), 6, 2))
                    
                    .TypeButtonText = Trim(ArrayRecords(0, lCount)) + Trim(ArrayRecords(1, lCount))    'YARD ADDR, PLATE_COUNT
                    .TypeButtonTextColor = vbBlack
'                    If VarType(ArrayRecords(3, lCount)) = vbNull Then
'                        .CellTag = ""
'                    Else
'                        .CellTag = Trim(ArrayRecords(3, lCount))  'ORD_NO
'                    End If
    
                    .Row = Val(sTemp) * 2
                    If VarType(ArrayRecords(2, lCount)) = vbNull Then
                      .TypeButtonText = ""  'LEN, WID
                    Else
                      .TypeButtonText = Trim(ArrayRecords(2, lCount))  'LEN, WID
                    End If
                    .TypeButtonTextColor = vbBlack
                End If
            Next lCount
            
        End If

        .ReDraw = True
        
    End With
    
    Only_Display = True
    Screen.MousePointer = vbDefault
    
    Exit Function
   
Error_Rtn:

    Set AdoRs = Nothing
    Only_Display = False
    MsgBox (Error)
    Screen.MousePointer = vbDefault
    
End Function

Public Function Addr_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn
    
    Dim lCount As Long
    Dim lCol As Long
    Dim lRow As Long
    
    Dim sTemp As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Addr_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With Sc.Item("Spread")

        Addr_Display = True
        .ReDraw = False
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("没有相关的数据", "I")
                
            Addr_Display = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = 0
            
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 2) > 0 Then  'Address
            
            For lCount = 0 To UBound(ArrayRecords, 2)
                
                If sTemp <> Mid(Trim(ArrayRecords(0, lCount)), 4, 2) Then
                    sTemp = Mid(Trim(ArrayRecords(0, lCount)), 4, 2)
                    
                    .BlockMode = True
                    .Row = Val(sTemp) * 2 - 1
                    .Row2 = Val(sTemp) * 2 - 1
                    .Col = 1: .Col2 = -1
                    .TypeButtonColor = &HF2F2F2
                    .TypeButtonTextColor = &H808080
                    .BlockMode = False
                
                    .BlockMode = True
                    .Row = Val(sTemp) * 2
                    .Row2 = Val(sTemp) * 2
                    .Col = 1: .Col2 = -1
                    .TypeButtonColor = &HA4FDE2
                    .TypeButtonTextColor = &H808080
                    .BlockMode = False
            
                End If
                
                .Row = Val(sTemp) * 2 - 1
                .Col = Val(Mid(Trim(ArrayRecords(0, lCount)), 6, 2))
                .TypeButtonText = Trim(ArrayRecords(0, lCount))   'YARD ADDR
'                .Text = Trim(ArrayRecords(0, lCount))

            Next lCount
            
        End If
        
        .ReDraw = True
        
    End With
    
    Addr_Display = True
    Screen.MousePointer = vbDefault
    
    Exit Function
   
Error_Rtn:

    Set AdoRs = Nothing
    Addr_Display = False
    Screen.MousePointer = vbDefault
    
End Function

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim sAddr As String
    
    ss1.Col = Col
    If Row Mod 2 = 1 Then
        ss1.Row = Row
        sAddr = Mid(ss1.TypeButtonText, 1, 7)
    Else
        ss1.Row = Row - 1
        sAddr = Mid(ss1.TypeButtonText, 1, 7)
    End If
    
    ss1.Row = ss1.Row + 1
'    If ss1.TypeButtonText = "" Then Exit Sub
    
    Load CGE2030C
    
    If ss1.TypeButtonColor = &HC0C0FF Then
        CGE2030C.txt_t_addr.Text = sAddr
    Else
        CGE2030C.txt_f_addr.Text = sAddr
    End If
    
    CGE2030C.text_cur_inv_code.Text = CBO_CUR_INV.Text
    CGE2030C.sOth = "CGE2040C"
    
    CGE2030C.Show
    CGE2030C.SetFocus
    'Unload Me
    
End Sub

Private Sub CBO_CUR_INV_Change()

    If Len(Trim(CBO_CUR_INV.Text)) = 2 Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", CBO_CUR_INV.Text, 2)
    Else
      text_cur_inv.Text = ""
    End If
    
End Sub

Private Sub CBO_CUR_INV_DblClick()
    Call CBO_CUR_INV_KeyUp(vbKeyF4, 0)
End Sub
Private Sub CBO_CUR_INV_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=CBO_CUR_INV
        DD.rControl.Add Item:=text_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
     
        If Len(Trim(CBO_CUR_INV.Text)) = 2 Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", CBO_CUR_INV.Text, 2)
        Else
          text_cur_inv.Text = ""
        End If
        
    End If
End Sub


