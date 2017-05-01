VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFT2010C 
   Caption         =   "生产上下限输入界面_AFT2010C"
   ClientHeight    =   9225
   ClientLeft      =   375
   ClientTop       =   2190
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
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
      ItemData        =   "AFT2010C.frx":0000
      Left            =   1515
      List            =   "AFT2010C.frx":000A
      TabIndex        =   1
      Tag             =   "SYSTEM"
      Top             =   120
      Width           =   1800
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   8625
      Left            =   60
      TabIndex        =   0
      Top             =   525
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15214
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
      MaxCols         =   18
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AFT2010C.frx":0023
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   225
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "炼钢/轧钢"
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
Attribute VB_Name = "AFT2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Template System
'-- Sub_System Name   Common
'-- Program Name      AGC2010C
'-- Program ID        AGC2010C
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
Public sQuery_load As String        'Active Form sQuery Setting
Public sQuery_Rt As String          'Active Form sQuery Setting

    
Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

 
Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection
 
Dim Proc_Sc As New Collection       'Spread Struc Collection
 
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    ' Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        
 Call Gp_Ms_Collection(CBO_PLT, "p", "n", " ", " ", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
   
     
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", "a", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     
     'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFT2010C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AFT2010C.P_REFER", Key:="P-R"
    sc1.Add Item:="AFT2010C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
     
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
   
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
   
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "F-System.INI", Me.Name)

    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing
    
    Set Mc1 = Nothing
  
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

   If Gf_Sp_Cls(sc1) Then
      Call Gp_Ms_Cls(Mc1("rControl"))
      Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
      Call Gp_Ms_ControlLock(Mc1("lControl"), False)
      CBO_PLT.Enabled = True
       pControl(1).SetFocus
   End If
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, sc1("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(sc1)
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(sc1)
  '  Call Gp_Sp_InAuthority(Sc1, 10)

End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(sc1)
    Call Gp_Sp_InAuthority(sc1, 10)
    
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

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(sc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Ref()
Dim iCnt As Integer
Dim PRC  As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("pControl"), Mc1("pControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.SetFocus
        ss1.OperationMode = OperationModeNormal
        CBO_PLT.Enabled = False
        
        With ss1
            For iCnt = 1 To .MaxRows
                .Row = iCnt
                .Col = 2
                Select Case Left(.Text, 2)
                    Case "BA"
                          PRC = "倒罐站"
                    Case "BB"
                          PRC = "铁水预处理"
                    Case "BC"
                          PRC = "转炉"
                    Case "BD"
                          PRC = "LF"
                    Case "BE"
                          PRC = "VD"
                    Case "BH"
                          PRC = "RH"
                    Case "BF"
                        If .Text = "BF21" Then
                            PRC = "连铸"
                        ElseIf .Text = "BF22" Then
                            PRC = "板坯切割"
                        End If
                        
                    Case "CA"
                        If .Text = "CA11" Then
                            PRC = "板卷加热炉"
                        ElseIf .Text = "CA12" Then
                            PRC = "中板加热炉"
                        ElseIf .Text = "CA13" Then
                            PRC = "加热炉缺号"
                        End If
                    
                    Case "CB"
                        If .Text = "CB12" Then
                            PRC = "轧钢"
                        ElseIf .Text = "CB13" Then
                            PRC = "母板分段"
                        End If
                    
                    Case "CC"
                        PRC = "GP_COILIF"
                    
                    Case "CE"
                        PRC = "GP_CBRSTIF"
                    
                    Case "CF"
                        PRC = "GP_DSSRSTIF"
                    
                    Case "CG"
                        If .Text = "CG11" Then
                            PRC = "钢板剪切"
                        ElseIf .Text = "CG12" Then
                            PRC = ""
                        End If
                        
                    Case "CJ"
                        PRC = ""
                    
             End Select
             
             .Col = 3
             .Text = PRC
             
             .Col = 4
             If .Value <> "1" Then
                .Col = 5
                .Text = ""
                .Lock = True
                .Col = 6
                .Text = ""
                .Lock = True
             End If
             
             .Col = 7
             If .Value <> 1 Then
                .Col = 8
                .Text = ""
                .Lock = True
                .Col = 9
                .Text = ""
                .Lock = True
             End If
             
             .Col = 10
             If .Value <> 1 Then
                .Col = 11
                .Text = ""
                .Lock = True
                .Col = 12
                .Text = ""
                .Lock = True
             End If
             
             .Col = 13
             If .Value <> 1 Then
                .Col = 14
                .Text = ""
                .Lock = True
                .Col = 15
                .Text = ""
                .Lock = True
             End If
             
             .Col = 16
             If .Value <> 1 Then
                .Col = 17
                .Text = ""
                .Lock = True
                .Col = 18
                .Text = ""
                .Lock = True
             End If
         Next
         
       End With
     
    End If
  
End Sub

'Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
'
'    lBlkcol1 = BlockCol
'    lBlkcol2 = BlockCol2
'    lBlkrow1 = BlockRow
'    lBlkrow2 = BlockRow2
'
'End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    Dim sStatus As String
 
    If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       sStatus = MDIMain.StatusBar1.Panels(1)
    
       Call Form_Ref
       MDIMain.StatusBar1.Panels(1) = sStatus
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

With ss1

    If Col = 4 Or Col = 7 Or Col = 10 Or Col = 13 Or Col = 16 Then
       
       .Col = .ActiveCol
       .Row = .ActiveRow
                
        If .Value = 0 Then
           .Text = "0"
           .Col = Col + 1
           .Lock = False
           
           .Col = Col + 2
           .Lock = False
           
        ElseIf .Value = 1 Then
           .Text = "1"
           .Col = Col + 1
           .Text = ""
           .Lock = True
           
           .Col = Col + 2
           .Text = ""
           .Lock = True
        End If
    
    End If
    
End With

End Sub

'Private Sub ss1_LostFocus()
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        
       ' Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    
    With ss1
             .Col = 4
             If .Value <> 1 Then
                .Col = 5
                .Text = ""
                .Lock = True
                .Col = 6
                .Text = ""
                .Lock = True
             End If
             
             .Col = 7
             If .Value <> 1 Then
                .Col = 8
                .Text = ""
                .Lock = True
                .Col = 9
                .Text = ""
                .Lock = True
             End If
             
             .Col = 10
             If .Value <> 1 Then
                .Col = 11
                .Text = ""
                .Lock = True
                .Col = 12
                .Text = ""
                .Lock = True
             End If
             
             .Col = 13
             If .Value <> 1 Then
                .Col = 14
                .Text = ""
                .Lock = True
                .Col = 15
                .Text = ""
                .Lock = True
             End If
             
             .Col = 16
             If .Value <> 1 Then
                .Col = 17
                .Text = ""
                .Lock = True
                .Col = 18
                .Text = ""
                .Lock = True
             End If
    End With
End Sub

'
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
