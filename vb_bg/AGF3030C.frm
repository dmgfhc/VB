VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGF3030C 
   Caption         =   "卷筒报废实绩查询及修改界面"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_ROLL_NO 
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
      ItemData        =   "AGF3030C.frx":0000
      Left            =   1230
      List            =   "AGF3030C.frx":0002
      TabIndex        =   0
      Tag             =   "ROLL_NO"
      Top             =   120
      Width           =   1365
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "卷筒号"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8610
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   15187
      _StockProps     =   64
      ColHeaderDisplay=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   499
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGF3030C.frx":0004
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3555
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "报废日期"
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
   Begin InDate.UDate SDT_TO_DATE 
      Height          =   315
      Left            =   6690
      TabIndex        =   2
      Tag             =   "终止日期"
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
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
   Begin InDate.UDate SDT_FROM_DATE 
      Height          =   315
      Left            =   4935
      TabIndex        =   3
      Tag             =   "起始日期"
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6450
      TabIndex        =   4
      Top             =   180
      Width           =   255
   End
End
Attribute VB_Name = "AGF3030C"
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
'-- Program Name      卷筒使用实绩查询及修改界面
'-- Program ID        AGF3020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHANG
'-- Coder             ZHANG
'-- Date              2009.10.10
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
Public sQuery_load As String        'Active Form sQuery Setting
Public QueryYN      As Boolean

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
 
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection




  Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Sheet"


    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDT_FROM_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDT_TO_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    
     
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "P", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  
  
  
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGF3030C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGF3030C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGF3030C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Call Gp_Sp_ColHidden(ss1, 2, True)

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub



'Private Sub chk_y_Click(Value As Integer)
'     Call Form_Ref
'End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

'Private Sub Form_Load()
'
'    Screen.MousePointer = vbHourglass
'
'    sAuthority = Gf_Pgm_Authority(Me.Name)
'
'    Call Form_Define
'
'    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
'
'    Call Gp_Ms_Cls(Mc1("rControl"))
'    Call Gp_Ms_NeceColor(Mc1("nControl"))
'
'    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
'    Call Gf_Sp_Cls(Proc_Sc("Sc"))
'    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
'
'
'    Screen.MousePointer = vbDefault
'
'    sQuery_load = "SELECT ROLL_NO FROM gp_roll WHERE  ROLL_NO LIKE  'J%' AND   ROLL_DISUSE_DATE IS NULL ORDER BY ROLL_NO "
'    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
'
'End Sub
Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
   
  Screen.MousePointer = vbDefault
    
    sQuery_load = "SELECT ROLL_NO FROM gp_roll WHERE  ROLL_NO LIKE  'J%'  ORDER BY ROLL_NO "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
'        Cancel = 1
'        Exit Sub
'    End If
'
'    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
'
'    Set pControl = Nothing
'    Set nControl = Nothing
'    Set iControl = Nothing
'    Set rControl = Nothing
'    Set cControl = Nothing
'    Set aControl = Nothing
'    Set lControl = Nothing
'    Set mControl = Nothing
'
'    Set iColumn1 = Nothing
'    Set pColumn1 = Nothing
'    Set lColumn1 = Nothing
'    Set nColumn1 = Nothing
'    Set mColumn1 = Nothing
'    Set aColumn1 = Nothing
'
'    Set Mc1 = Nothing
'    Set sc1 = Nothing
'    Set Proc_Sc = Nothing
'
'    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
'
'End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)

    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Spread_Can()
Dim sCid As String
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    ss1.Row = ss1.ActiveRow
    ss1.Col = 2
    sCid = ss1.Text
    If sCid <> "" Then
       ss1.Col = 1
       ss1.Text = sCid
    End If
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        SDT_FROM_DATE.SetFocus
        SDT_FROM_DATE = ""
        SDT_TO_DATE = ""

     
    End If

End Sub

Public Sub Form_Ref()

Dim iRow  As Integer
Dim I, j, Scr_wgt, Hm_wgt, Steel_wgt As Integer
Dim Iron_Rec, Iron_Use, Back_Wgt As Double
On Error GoTo Refer_Err
Dim sCid As String

QueryYN = False


    If SDT_FROM_DATE.RawData = "" And Trim(CBO_ROLL_NO.Text) = "" Then
       SDT_FROM_DATE.Text = Format(Now, "YYYY-MM") + "-01"
    End If
    If SDT_TO_DATE.RawData = "" And Trim(CBO_ROLL_NO.Text) = "" Then
       SDT_TO_DATE.Text = Format(Now, "YYYY-MM-DD")
    End If


 If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

      If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
         Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

         Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
        
                For I = 1 To ss1.MaxRows
                    ss1.Col = 2
                    ss1.Row = I
                   sCid = ss1.Text
                   If sCid <> "" Then
                   ss1.Col = 1
                   ss1.Text = sCid
                  End If
                Next I
        
           With ss1

            For I = 1 To .MaxRows
                  .Row = I
                  .Col = 4
                  Iron_Rec = Iron_Rec + Val(.Text)

             Next I
               .MaxRows = .MaxRows + 1
               .Row = .MaxRows
             For I = 1 To .MaxCols
                   .Col = I
                   .BackColor = "&HE6E6FF"
             Next I

             .Col = 1
             .Text = "合计"
             .Lock = True
             .Col = 3
             .Lock = True
             .Col = 4
             .Text = Str(Iron_Rec)
             .Lock = True

        End With
    End If
 Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
   Dim icount As Integer
   Dim MsgBox As String
   

     ss1.Row = ss1.ActiveRow
     ss1.Col = 2
     MsgBox = "您确定要报废这个卷筒号" + ss1.Text + "吗？"
    
     If Not Gf_MessConfirm(MsgBox, "Q") Then Exit Sub


   For icount = 1 To ss1.MaxRows

        Select Case Trim(Gf_Sp_RcvData(ss1, 0, icount))

           Case "Input", "Update"

             With ss1
             .Col = 3
             If Not Gp_DateCheck(.Text, "S") Then
                         Call Gp_MsgBoxDisplay("请输入正确的使用时间")
                         Exit Sub
                      End If
                  End With
        End Select

    Next icount


   If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
        Call Form_Ref
   End If

End Sub



Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.Col = 12
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID
    
    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 1
    ss1.BackColor = &HC0FFFF
    
       Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT ROLL_NO  FROM GP_ROLL WHERE ROLL_NO LIKE 'J%' ORDER BY ROLL_NO ")
       
       
    
End Sub
Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
    
End Sub
Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    ss1.Col = 12
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID
    
    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 1
    ss1.BackColor = &HC0FFFF
    
       Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT ROLL_NO  FROM GP_ROLL WHERE ROLL_NO LIKE 'J%' ORDER BY ROLL_NO ")
    ss1.OperationMode = OperationModeNormal

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
Dim sCid As String
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Private Sub SDT_FROM_DATE_DblClick()
    If SDT_FROM_DATE.RawData = "" Then
     SDT_FROM_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_TO_DATE.RawData = "" Then
        SDT_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub
Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim iRow  As Integer
Dim I, j, Scr_wgt, Hm_wgt, Steel_wgt As Integer

   If Row <> 0 Then
    If Col = 3 Then
        ss1.Col = Col
        ss1.Row = Row
        If ss1.Lock = False Then
           ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
           ss1.Col = 0
       
       
           If ss1.Text <> "Input" And ss1.Text <> "Delete" Then
              ss1.Text = "Update"
           End If
           

        End If
        
          
End If

    End If
End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim sCid As String
If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        ss1.Col = 0
        ss1.Row = ss1.ActiveRow
        If ss1.Text = "Update" Then
            ss1.Col = 12
            ss1.Text = sUserID
        End If
        
        If Col = 1 Then
           ss1.Col = 1
           sCid = ss1.Text
           ss1.Col = 2
           ss1.Text = sCid
        End If
    End If
    
End Sub
Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
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
Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

  If ss1.ActiveCol = 1 Then
       ss1.Row = ss1.ActiveRow
       ss1.Col = ss1.ActiveCol
       If Len(Trim(ss1.Text)) = 7 Then
          Dim sQuery As String
          sQuery = "SELECT SUBSTR(ROLL_IN_DATE,1,4) || '_'|| SUBSTR(ROLL_IN_DATE,5,2) || '_'|| SUBSTR(ROLL_IN_DATE,7,2)|| '_'|| SUBSTR(ROLL_IN_DATE,9,2)|| '_'|| SUBSTR(ROLL_IN_DATE,11,2)||'_'|| SUBSTR(ROLL_IN_DATE,13,2)FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 5
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
    
          
          ss1.Col = 1
          sQuery = "SELECT ROLL_WGT FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 6
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
          ss1.Col = 1
          sQuery = "SELECT  ISSUETALLYNO   FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 7
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          ss1.Col = 1
          sQuery = "SELECT MTRLNO FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 8
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          
          ss1.Col = 1
          sQuery = "SELECT ROLL_PRICE FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 9
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
          
       Else
       
          ss1.Col = 5
          ss1.Text = ""
       
          ss1.Col = 6
          ss1.Text = ""
          
         
          ss1.Col = 7
          ss1.Text = ""
          
      
          ss1.Col = 8
          ss1.Text = ""
          
      
          ss1.Col = 9
          ss1.Text = ""
          
    End If
  End If
End Sub
Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)


  If ss1.ActiveCol = 1 Then
       ss1.Row = ss1.ActiveRow
       ss1.Col = ss1.ActiveCol
       If Len(Trim(ss1.Text)) = 7 Then
          Dim sQuery As String
          sQuery = "SELECT SUBSTR(ROLL_IN_DATE,1,4) || '_'|| SUBSTR(ROLL_IN_DATE,5,2) || '_'|| SUBSTR(ROLL_IN_DATE,7,2)|| '_'|| SUBSTR(ROLL_IN_DATE,9,2)|| '_'|| SUBSTR(ROLL_IN_DATE,11,2)||'_'|| SUBSTR(ROLL_IN_DATE,13,2)FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 5
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
    
          
          ss1.Col = 1
          sQuery = "SELECT ROLL_WGT FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 6
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
          ss1.Col = 1
          sQuery = "SELECT  ISSUETALLYNO   FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 7
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          ss1.Col = 1
          sQuery = "SELECT MTRLNO FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 8
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          
          ss1.Col = 1
          sQuery = "SELECT ROLL_PRICE FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 9
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
          
       Else
       
          ss1.Col = 5
          ss1.Text = ""
       
          ss1.Col = 6
          ss1.Text = ""
          
         
          ss1.Col = 7
          ss1.Text = ""
          
      
          ss1.Col = 8
          ss1.Text = ""
          
      
          ss1.Col = 9
          ss1.Text = ""
          
    End If
  End If
End Sub
Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Public Function Pf_ComboAdd(Conn As ADODB.Connection, ss As vaSpread, Col As Integer, sQuery As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim AdoRs As ADODB.Recordset
    Dim sList As String
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Pf_ComboAdd = False: Exit Function
    End If
    
'    If ClsChk Then
'        Cbo.Clear
'    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If AdoRs.Fields(0) <> vbNull Then
                sList = sList & AdoRs.Fields(0) & vbTab
                'Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Pf_ComboAdd = True
    Else
        Pf_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    ss.Col = Col
    ss.TypeComboBoxList = sList
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Pf_ComboAdd = False

End Function

