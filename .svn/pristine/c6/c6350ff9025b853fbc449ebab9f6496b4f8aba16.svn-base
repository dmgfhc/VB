VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AKW2010C 
   Caption         =   "二级系统状态查询_AKW2010C"
   ClientHeight    =   8985
   ClientLeft      =   255
   ClientTop       =   1125
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   14610
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   465
      Top             =   390
   End
   Begin InDate.ULabel ULabel50 
      Height          =   345
      Left            =   10590
      Top             =   420
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   609
      Caption         =   ""
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
      Height          =   7785
      Left            =   1530
      TabIndex        =   0
      Top             =   1080
      Width           =   12960
      _Version        =   393216
      _ExtentX        =   22860
      _ExtentY        =   13732
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
      MaxCols         =   4
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKW2010C.frx":0000
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   5250
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   4860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   4455
      Width           =   435
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   4050
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   1650
      Width           =   435
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   2052
      Width           =   435
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   2454
      Width           =   435
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   2856
      Width           =   435
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   3258
      Width           =   435
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000007&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1095
      Shape           =   3  'Circle
      Top             =   3660
      Width           =   435
   End
End
Attribute VB_Name = "AKW2010C"
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
'-- Program Name      L2 And L3 Connection Normal Or Abnormal
'-- Program ID        AKW2010C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2004.8.23
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
Public iCDS, iBOF, iMSP, iCCM, iRHF, iMILL, iCDS2, iBOF2, iMSP2, iCCM2, iCDS3, iBOF3, iCCM3 As Long

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKW2010C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
            
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Call Gp_Sp_ColHidden(ss1, 5, True)
        
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

    Dim sQuery As String

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "K-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    sQuery = Gf_FloatFind(M_CN1, "SELECT SYSDATE FROM DUAL")
    ULabel50.Caption = "查询时间:  " + sQuery
    
    If MDIMain.MenuTool.Buttons(2).Enabled = False Then
        Call ColumNameSet

        Timer1.Enabled = False
        Exit Sub
    Else
       Call Form_Ref
       Timer1.Enabled = True
        With ss1
            .Enabled = True
            .Col = 3
            .ROW = 1
            iCDS = CInt(.Text)
            .ROW = 2
            iBOF = CInt(.Text)
            .ROW = 3
            iMSP = CInt(.Text)
            .ROW = 4
            iCCM = CInt(.Text)
            .ROW = 5
            iRHF = CInt(.Text)
            .ROW = 6
            iMILL = CInt(.Text)
            .ROW = 7
            iCCM2 = CInt(.Text)
            .ROW = 8
            iMSP2 = CInt(.Text)
            
            .RowHeight(1) = 18
            .RowHeight(2) = 18
            .RowHeight(3) = 18
            .RowHeight(4) = 18
            .RowHeight(5) = 18
            .RowHeight(6) = 18
            .RowHeight(7) = 18
            .RowHeight(8) = 18
        End With
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "K-System.INI", Me.Name)
        
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        ULabel50.Caption = ""
        Call ColumNameSet
        Timer1.Enabled = False
    End If

End Sub

Public Sub Form_Ref()

    Dim iRow As Integer
    Dim sQuery As String
        
    sQuery = "SELECT    '正常', LINK_TIME, WATCH_DOG,"
    sQuery = sQuery & " (TO_DATE(SUBSTR(TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS'),1,12),'YYYYMMDDHH24MISS') - TO_DATE(SUBSTR(LINK_TIME,1,12),'YYYYMMDDHH24MISS')) * 24 * 60"
    sQuery = sQuery & " FROM FP_WATCHDOGIF"
    sQuery = sQuery & " ORDER BY PRC_LINE, LINK_ID"
    
    If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       
        With ss1
            Call ColumNameSet
            
            sQuery = Gf_FloatFind(M_CN1, "SELECT SYSDATE FROM DUAL")
            ULabel50.Caption = "查询时间:  " + sQuery
            
            .Col = 2
             For iRow = 1 To .MaxRows
                .ROW = iRow
                If .Text <> "" Then
                   .Text = Mid(.Text, 1, 4) + "-" + Mid(.Text, 5, 2) + "-" + Mid(.Text, 7, 2) + " " + _
                           Mid(.Text, 9, 2) + ":" + Mid(.Text, 11, 2) + ":" + Mid(.Text, 13, 2)
                End If
             Next iRow
            
        End With
        
        If Timer1.Enabled = False Then
           Timer1.Enabled = True
        End If
    
    End If

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub Timer1_Timer()

    Dim sColor As String
    Call Form_Ref
    
    With ss1
         .Col = 3
         .ROW = 1
         If CInt(.Text) = iCDS Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape1.FillColor = &HFF
            Shape1.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape1.FillColor = &HFF00&
            Shape1.BorderColor = &HFF00&
         End If
         
         .Col = 3
         .ROW = 2
         If CInt(.Text) = iBOF Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape2.FillColor = &HFF
            Shape2.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape2.FillColor = &HFF00&
            Shape2.BorderColor = &HFF00&
         End If
         
         .Col = 3
         .ROW = 3
         If CInt(.Text) = iMSP Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape3.FillColor = &HFF
            Shape3.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape3.FillColor = &HFF00&
            Shape3.BorderColor = &HFF00&
         End If
         
         .Col = 3
         .ROW = 4
         If CInt(.Text) = iCCM Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape4.FillColor = &HFF
            Shape4.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape4.FillColor = &HFF00&
            Shape4.BorderColor = &HFF00&
         End If
         
         .Col = 3
         .ROW = 5
         If CInt(.Text) = iRHF Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape5.FillColor = &HFF
            Shape5.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape5.FillColor = &HFF00&
            Shape5.BorderColor = &HFF00&
         End If
         
         .Col = 3
         .ROW = 6
         If CInt(.Text) = iMILL Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape6.FillColor = &HFF
            Shape6.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape6.FillColor = &HFF00&
            Shape6.BorderColor = &HFF00&
         End If
         
         .Col = 3
         .ROW = 7
         If CInt(.Text) = iCCM2 Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape7.FillColor = &HFF
            Shape7.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape7.FillColor = &HFF00&
            Shape7.BorderColor = &HFF00&
         End If
         
         .Col = 3
         .ROW = 8
         If CInt(.Text) = iMSP2 Then
            .Col = 1
            .Text = "非正常"
            .BackColor = &HFF
            Shape8.FillColor = &HFF
            Shape8.BorderColor = &HFF
         Else
            .Col = 1
            .Text = "正常"
            .Col = 2
            sColor = .BackColor
            .BackColor = sColor
            Shape8.FillColor = &HFF00&
            Shape8.BorderColor = &HFF00&
         End If
         
'         .Col = 3
'         .ROW = 9
'         If CInt(.Text) = iMSP2 Then
'            .Col = 1
'            .Text = "非正常"
'            .BackColor = &HFF
'            Shape9.FillColor = &HFF
'            Shape9.BorderColor = &HFF
'         Else
'            .Col = 1
'            .Text = "正常"
'            .Col = 2
'            sColor = .BackColor
'            .BackColor = sColor
'            Shape9.FillColor = &HFF00&
'            Shape9.BorderColor = &HFF00&
'         End If
'
'         .Col = 3
'         .ROW = 10
'         If CInt(.Text) = iCCM2 Then
'            .Col = 1
'            .Text = "非正常"
'            .BackColor = &HFF
'            Shape10.FillColor = &HFF
'            Shape10.BorderColor = &HFF
'         Else
'            .Col = 1
'            .Text = "正常"
'            .Col = 2
'            sColor = .BackColor
'            .BackColor = sColor
'            Shape10.FillColor = &HFF00&
'            Shape10.BorderColor = &HFF00&
'         End If
         
        .Col = 3
        .ROW = 1
        iCDS = CInt(.Text)
        .ROW = 2
        iBOF = CInt(.Text)
        .ROW = 3
        iMSP = CInt(.Text)
        .ROW = 4
        iCCM = CInt(.Text)
        .ROW = 5
        iRHF = CInt(.Text)
        .ROW = 6
        iMILL = CInt(.Text)
        .ROW = 7
        iCCM2 = CInt(.Text)
        .ROW = 8
        iMSP2 = CInt(.Text)
         
    End With
    
End Sub

Public Sub ColumNameSet()

    Dim iRow As Integer
    
    With ss1
        .OperationMode = OperationModeNormal
        
        .RowHeight(0) = 24
        
        .ColWidth(0) = 20
                        
        .MaxRows = 8
        .Col = 0
        .ROW = 0
        .Text = "二级系统分类"
        .ROW = 1
        .Text = "CDS-铁水预处理(#1)"
        .FontBold = True
        .ROW = 2
        .Text = "BOF-转炉"
        .FontBold = True
        .ROW = 3
        .Text = "LF/VD-精炼(#1)"
        .FontBold = True
        .ROW = 4
        .Text = "CCM-连铸(#1)"
        .FontBold = True
        .ROW = 5
        .Text = "RHF-加热炉"
        .FontBold = True
        .ROW = 6
        .Text = "MILL-轧钢"
        .FontBold = True
        .ROW = 7
        .Text = "CCM-连铸(#2)"
        .FontBold = True
        .ROW = 8
        .Text = "LF-精炼(#2)"
        .FontBold = True
        
        .Col = 1
         For iRow = 1 To .MaxRows
            .ROW = iRow
            .Text = "正常"
         Next iRow
         
        Shape1.FillColor = &HFF00&
        Shape1.BorderColor = &HFF00&
        Shape2.FillColor = &HFF00&
        Shape2.BorderColor = &HFF00&
        Shape3.FillColor = &HFF00&
        Shape3.BorderColor = &HFF00&
        Shape4.FillColor = &HFF00&
        Shape4.BorderColor = &HFF00&
        Shape5.FillColor = &HFF00&
        Shape5.BorderColor = &HFF00&
        Shape6.FillColor = &HFF00&
        Shape6.BorderColor = &HFF00&
        Shape7.FillColor = &HFF00&
        Shape7.BorderColor = &HFF00&
        Shape8.FillColor = &HFF00&
        Shape8.BorderColor = &HFF00&
'        Shape9.FillColor = &HFF00&
'        Shape9.BorderColor = &HFF00&
'        Shape10.FillColor = &HFF00&
'        Shape10.BorderColor = &HFF00&
        
        .RowHeight(1) = 18
        .RowHeight(2) = 18
        .RowHeight(3) = 18
        .RowHeight(4) = 18
        .RowHeight(5) = 18
        .RowHeight(6) = 18
        .RowHeight(7) = 18
        .RowHeight(8) = 18
'        .RowHeight(9) = 18
'        .RowHeight(10) = 18
                    
    End With
        
End Sub

