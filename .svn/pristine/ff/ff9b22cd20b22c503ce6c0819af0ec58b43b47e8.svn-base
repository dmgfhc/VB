VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFO2030C 
   Caption         =   "炼钢生产线停机实绩界面_AFO2030C"
   ClientHeight    =   9225
   ClientLeft      =   135
   ClientTop       =   2265
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_PLT 
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
      ItemData        =   "AFO2030C.frx":0000
      Left            =   1575
      List            =   "AFO2030C.frx":0007
      TabIndex        =   0
      Tag             =   "工厂代码"
      Top             =   135
      Width           =   1020
   End
   Begin VB.TextBox txt_PRC 
      Alignment       =   2  'Center
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
      Left            =   4635
      TabIndex        =   1
      Tag             =   "工序代码"
      Top             =   135
      Width           =   825
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8655
      Left            =   45
      TabIndex        =   2
      Top             =   510
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   15266
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
      MaxCols         =   16
      MaxRows         =   2
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AFO2030C.frx":000F
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
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
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3240
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "工序"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   6165
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "发生时间"
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
   Begin InDate.UDate txt_OCCR_TS 
      Height          =   315
      Left            =   7590
      TabIndex        =   5
      Top             =   135
      Width           =   1455
      _ExtentX        =   2566
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
   Begin InDate.UDate txt_OCCR_TS2 
      Height          =   315
      Left            =   9270
      TabIndex        =   4
      Top             =   135
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9075
      TabIndex        =   3
      Top             =   180
      Width           =   255
   End
End
Attribute VB_Name = "AFO2030C"
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
'-- Program Name      DELAY
'-- Program ID        AFO2030C
'-- Designer          ZHENG WEN
'-- Coder             ZHENG WEN
'-- Date              2003.8.6
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
Public sProc_cd As String

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
       Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PRC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_OCCR_TS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_OCCR_TS2, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'PRC_LINE
    Call Gp_Sp_Collection(ss1, 3, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFO2030C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AFO2030C.P_REFER", Key:="P-R"
    sc1.Add Item:="AFO2030C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 5)
    
    CBO_PLT.Text = "B1"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
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
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Exit Sub
    End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
Dim Time1, TiME2 As String
Dim i As Integer
With ss1
     For i = 1 To .MaxRows
        .Row = i
        .Col = 0
        If .Text = "Update" Or .Text = "Input" Then
           .Col = 7
            If Trim(Mid(.Text, 1, 4)) = "" Or Trim(.Text) = "" Then
               MsgBox "停机开始时间必须输入！", vbCritical, "系统提示信息"
               Exit Sub
            End If
            
            Time1 = .Text
           .Col = 10
            TiME2 = .Text
            If (TiME2 < Time1) And Trim(TiME2) <> "    -  -     :  :  " And Trim(TiME2) <> "" Then
               MsgBox "停机结束时间应该晚于停机开始时间！", vbCritical, "系统提示"
               Exit Sub
            End If
        End If
     Next i
End With
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.Row = ss1.ActiveRow
    ss1.Col = 1
    ss1.Text = Trim(CBO_PLT.Text)
 
    ss1.Col = 15
    ss1.Text = sUserID
    
    ss1.Col = 5
    Call ss1_EditChange(ss1.Col, ss1.ActiveRow)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 10
    ss1.Text = ""
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 15
    ss1.Text = sUserID
    ss1.Col = 4
    Call ss1_EditChange(ss1.Col, ss1.ActiveRow)
    
    ss1.Col = 5
    Call ss1_EditChange(ss1.Col, ss1.ActiveRow)
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
        
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

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


Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
Dim v_occr_date, v_shift As String
If Col = 4 Then
   ss1.Row = Row
   ss1.Col = Col
   v_occr_date = Mid(ss1.Text, 12, 2) + Mid(ss1.Text, 15, 2) + Mid(ss1.Text, 18, 2)
   If v_occr_date >= "000000" And v_occr_date < "080000" Then
      v_shift = "1"
   ElseIf v_occr_date >= "080000" And v_occr_date < "160000" Then
      v_shift = "2"
   ElseIf v_occr_date >= "160000" And v_occr_date < "240000" Then
      v_shift = "3"
   End If
        
   ss1.Col = 12
   ss1.Col2 = 12
   ss1.Row = ss1.ActiveRow
   ss1.Text = v_shift
   
ElseIf Col = 10 Then
   ss1.Row = Row
   ss1.Col = Col
   If Trim(ss1.Text) = "-  -     :  :" Then
      ss1.Text = ""
   End If

End If
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
If Col = 4 Then
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End If

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim v_occr_date, v_shift As String

If Row <> 0 Then
    If Col = 4 Or Col = 7 Or Col = 10 Then
        ss1.Col = Col
        ss1.Row = Row
        If ss1.Lock = False Then
           ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
           ss1.Col = 0
           If ss1.Text <> "Input" And ss1.Text <> "Delete" Then
              ss1.Text = "Update"
           End If
        End If
        
        If Col = 4 Then
           ss1.Col = Col
           ss1.Row = Row
           v_occr_date = Mid(ss1.Text, 12, 2) + Mid(ss1.Text, 15, 2) + Mid(ss1.Text, 18, 2)
           If v_occr_date >= "000000" And v_occr_date < "080000" Then
              v_shift = "1"
           ElseIf v_occr_date >= "080000" And v_occr_date < "160000" Then
              v_shift = "2"
           ElseIf v_occr_date >= "160000" And v_occr_date < "240000" Then
              v_shift = "3"
           End If
        
           ss1.Col = 12
           ss1.Row = ss1.ActiveRow
           ss1.Text = v_shift
        End If
    
    End If
    
    
End If
 
End Sub

Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim sTemp_Mana_Code As String
Dim sTemp_Code As String
Dim v_occr_date, v_shift As String
    ss1.Row = ss1.ActiveRow
    ss1.Col = 2
    sProc_cd = ss1.Text
    ss1.Col = Col
    If ss1.Col = 5 Then
       If ss1.Text = "" Then
          ss1.Col = 6
          ss1.Text = ""
       ElseIf Len(Trim(ss1.Text)) = 5 Then
          
          Select Case sProc_cd
                        Case "BB"
                          sTemp_Mana_Code = "F0010"
                        Case "BC"
                          sTemp_Mana_Code = "F0020"
                        Case "BD"
                          sTemp_Mana_Code = "F0021"
                        Case "BE"
                          sTemp_Mana_Code = "F0021"
                        Case "BF"
                          sTemp_Mana_Code = "F0022"
                        ''ADD BY GUOLI AT 200701081031''
                        Case "BG"
                          sTemp_Mana_Code = "F0028"
                        Case "BH"
                          sTemp_Mana_Code = "F0029"
                        ''''''''''''''''''''''''''''''''
                 End Select
          sTemp_Code = ss1.Text
          ss1.Col = 6
          ss1.Text = Gf_ComnNameFind(M_CN1, sTemp_Mana_Code, Trim(sTemp_Code), 1)
       End If
    End If
    
    If Col = 4 Then
        ss1.Row = Row
        ss1.Col = Col
        v_occr_date = Mid(ss1.Text, 12, 2) + Mid(ss1.Text, 15, 2) + Mid(ss1.Text, 18, 2)
        If v_occr_date >= "000000" And v_occr_date < "080000" Then
           v_shift = "1"
        ElseIf v_occr_date >= "080000" And v_occr_date < "160000" Then
           v_shift = "2"
        ElseIf v_occr_date >= "160000" And v_occr_date < "240000" Then
           v_shift = "3"
        End If
             
        ss1.Col = 12
        ss1.Col2 = 12
        ss1.Row = ss1.ActiveRow
        ss1.Text = v_shift
    End If

    If ss1.Col = 2 Then
       If Len(Trim(ss1.Text)) = 2 Then
          If ss1.Text <> "BA" And ss1.Text <> "BB" And ss1.Text <> "BC" And ss1.Text <> "BD" And ss1.Text <> "BE" And ss1.Text <> "BF" And ss1.Text <> "BG" And ss1.Text <> "BH" Then
             MsgBox "工序代码不正确！", vbCritical, "系统提示信息"
             Exit Sub
          End If
       End If
    End If
    
    If ss1.Col = 12 Then
       If Len(Trim(ss1.Text)) = 1 Then
          If ss1.Text <> "1" And ss1.Text <> "2" And ss1.Text <> "3" And ss1.Text <> "4" Then
             MsgBox "班次不正确！", vbCritical, "系统提示信息"
             Exit Sub
          End If
       End If
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

'    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_OCCR_TS_DblClick()
              
    txt_OCCR_TS.RawData = Format(Now, "YYYYMMDD")

End Sub

Private Sub SS1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    If ss1.ActiveCol = 2 Then

       If KeyCode = vbKeyF4 Then
         Set DD.sPname = ss1
    
         DD.sWitch = "SP"
         DD.sKey = "C0002"
         DD.rControl.Add Item:=2
        
         DD.nameType = "2"
        
         Call Gf_Common_DD(M_CN1, KeyCode)
        
         Exit Sub
        
       End If
    End If
        
    ss1.Col = 2
    ss1.Row = ss1.ActiveRow
    sProc_cd = ss1.Text
        
    If ss1.ActiveCol = 5 Then

       If KeyCode = vbKeyF4 Then
         
             If sProc_cd = "" Then
                 MsgBox "请先输入工序代码", vbCritical, "系统提示信息"
                 ss1.Col = 2
                 ss1.Row = ss1.ActiveRow
                 ss1.SetFocus
                 Call ss1.SetSelection(2, ss1.ActiveRow, 2, ss1.ActiveRow)
                 
                 Exit Sub
             Else
                 ss1.Col = 5
                 Set DD.sPname = ss1
            
                 DD.sWitch = "SP"
                 Select Case sProc_cd
                        Case "BB"
                          DD.sKey = "F0010"
                        Case "BC"
                          DD.sKey = "F0020"
                        Case "BD"
                          DD.sKey = "F0021"
                        Case "BE"
                          DD.sKey = "F0021"
                        Case "BF"
                          DD.sKey = "F0022"
                        Case "BG"
                          DD.sKey = "F0028"
                        Case "BH"
                          DD.sKey = "F0029"
                 End Select
                 
                 DD.rControl.Add Item:=5
                 DD.rControl.Add Item:=6
                
                 DD.nameType = "2"
                
                 Call Gf_Common_DD(M_CN1, KeyCode)
                
                 Exit Sub
             End If
            
       Else
           ss1.Col = ss1.ActiveCol
           
           If Len(Trim(ss1.Text)) = ss1.TypeMaxEditLen Then
              
              ss1.Col = 5
              sTemp_Code = ss1.Text
              
              Select Case sProc_cd
                     Case "BB"
                       DD.sKey = "F0010"
                     Case "BC"
                       DD.sKey = "F0020"
                     Case "BD"
                       DD.sKey = "F0021"
                     Case "BE"
                       DD.sKey = "F0021"
                     Case "BF"
                       DD.sKey = "F0022"
                     Case "BG"
                       DD.sKey = "F0028"
                     Case "BH"
                       DD.sKey = "F0029"
              End Select
              
              ss1.Col = 6
              ss1.Text = Gf_ComnNameFind(M_CN1, DD.sKey, Trim(sTemp_Code), 2)
              ss1.SetFocus
           Else
              ss1.Col = 6
              ss1.Text = ""
              ss1.SetFocus
           End If
       End If
    End If

End Sub

Private Sub txt_OCCR_TS2_DblClick()
    txt_OCCR_TS2.RawData = Format(Now, "YYYYMMDD")
End Sub

Private Sub txt_PRC_KeyUp(KeyCode As Integer, Shift As Integer)
       
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        DD.rControl.Add Item:=txt_PRC
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
     
    End If
       
End Sub
