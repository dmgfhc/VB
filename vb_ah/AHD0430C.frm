VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AHD0430C 
   Caption         =   "�����������ƽ�ⱨ��_AHD0430C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   13065
      TabIndex        =   11
      Top             =   180
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.OptionButton Option2 
      Caption         =   "���"
      Height          =   345
      Left            =   11715
      TabIndex        =   10
      Top             =   105
      Width           =   1080
   End
   Begin VB.OptionButton Option1 
      Caption         =   "����"
      Height          =   315
      Left            =   10575
      TabIndex        =   9
      Top             =   135
      Width           =   1005
   End
   Begin VB.TextBox txt_plt_name 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   7740
      MaxLength       =   40
      TabIndex        =   8
      Tag             =   "mill_plt"
      Top             =   105
      Width           =   2505
   End
   Begin VB.TextBox txt_plt 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   7230
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "plt"
      Top             =   105
      Width           =   495
   End
   Begin VB.TextBox TXT_CHECK_YARD 
      Height          =   300
      Left            =   7245
      TabIndex        =   6
      Top             =   105
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox TXT_CHECK_DATE 
      Height          =   285
      Left            =   9915
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox text_cur_inv_code 
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
      Left            =   4260
      MaxLength       =   2
      TabIndex        =   4
      Top             =   105
      Width           =   375
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   4650
      TabIndex        =   3
      Top             =   105
      Width           =   1050
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "��ӡ����"
      Height          =   330
      Left            =   13650
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   1260
   End
   Begin InDate.UDate dtp_yy_mm 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "����"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "����"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8925
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   15743
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   21
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AHD0430C.frx":0000
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3030
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "�ֿ�"
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
      ForeColor       =   -2147483646
   End
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   14
      Left            =   5835
      Top             =   105
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "AHD0430C"
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
'-- Program Name      Refer Template
'-- Program ID        Refer
'-- Document No       Q-00-0010(Specification)
'-- Designer          zheng wen
'-- Coder             zheng wen
'-- Date              2005.9.28
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

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2
Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(dtp_yy_mm, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(Text1, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AHD0430C.P_REFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Duplicate Count
    iDupCnt = 2
    
        'Sum Column Count
    iSumCnt = 17
    
    'Sum Column Setting
    iSumCol.Add Item:=21
    iSumCol.Add Item:=5
    iSumCol.Add Item:=6
    iSumCol.Add Item:=7
    iSumCol.Add Item:=8
    iSumCol.Add Item:=9
    iSumCol.Add Item:=10
    iSumCol.Add Item:=11
    iSumCol.Add Item:=12
    iSumCol.Add Item:=13
    iSumCol.Add Item:=14
    iSumCol.Add Item:=15
    iSumCol.Add Item:=16
    iSumCol.Add Item:=17
    iSumCol.Add Item:=18
    iSumCol.Add Item:=19
    iSumCol.Add Item:=20
'    iSumCol.Add Item:=21
'    iSumCol.Add Item:=22
    
    Me.KeyPreview = True

End Sub



Private Sub cmdReport_Click()
    If dtp_yy_mm.RawData = "" Then
       MsgBox "����������Ҫ��ѯ�����ڣ�", vbCritical, "ϵͳ��ʾ��Ϣ"
       Exit Sub
    End If
    
    If text_cur_inv_code = "" Then
       MsgBox "����������Ҫ��ѯ�Ĳֿ⣡", vbCritical, "ϵͳ��ʾ��Ϣ"
       Exit Sub
    
    End If

    Screen.MousePointer = vbHourglass
    
'    Call ExcelPrn
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call subButtonHide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
'    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
   
'    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    
    dtp_yy_mm.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")

    Screen.MousePointer = vbDefault
    
    Option1.Value = True
    Text1.Text = "1"

    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    Set iSumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subButtonHide
    End If
    Option1.Value = True
    Text1.Text = "1"

End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
    
    If dtp_yy_mm.RawData = "" Then
       MsgBox "����������Ҫ��ѯ�����ڣ�", vbCritical, "ϵͳ��ʾ��Ϣ"
       Exit Sub
    End If
    
    If text_cur_inv_code = "" Then
       MsgBox "����������Ҫ��ѯ�Ĳֿ⣡", vbCritical, "ϵͳ��ʾ��Ϣ"
       Exit Sub
    
    End If
    
    sQuery = "{call AHD0430C.P_REFER ('" + dtp_yy_mm.RawData + "','" + text_cur_inv_code + "','" + txt_plt.Text + "','" + Text1.Text + "')}"
    If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, 1, iDupCnt, iSumCnt, iSumCol, False) Then
'    If Gf_Sub_total_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
'        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call subButtonHide
        TXT_CHECK_DATE = dtp_yy_mm.RawData
        TXT_CHECK_YARD = text_cur_inv_code
    End If
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

'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0

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
Private Sub text_cur_inv_code_Change()
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
    Else
      text_cur_inv.Text = ""
    End If

End Sub
Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
End Sub

'Private Sub ExcelPrn()
'    Dim i               As Integer
'    Dim xlApp           As Object
'    Dim xlSheet         As Object
'    Dim sRow            As String
'
'    If ss1.MaxRows < 1 Then
'       MsgBox "���Ȳ�ѯ�����ٴ�ӡ��", vbCritical, "ϵͳ��ʾ��Ϣ"
'       Exit Sub
'    End If
'
'    If Trim(TXT_CHECK_DATE) <> dtp_yy_mm.RawData Then
'       MsgBox "ѡ�������û�н��в�ѯ�����Ȳ�ѯ�����ٴ�ӡ��", vbCritical, "ϵͳ��ʾ��Ϣ"
'       Exit Sub
'
'    End If
'
'    If Trim(TXT_CHECK_YARD) <> Trim(text_cur_inv_code) Then
'       MsgBox "ѡ��Ĳֿ�û�н��в�ѯ�����Ȳ�ѯ�����ٴ�ӡ��", vbCritical, "ϵͳ��ʾ��Ϣ"
'       Exit Sub
'
'    End If
'
'    Screen.MousePointer = vbHourglass
'
'    On Error Resume Next
'
'    Set xlApp = GetObject(, "Excel.Application")
'    If Err.Number <> 0 Then
'        Set xlApp = CreateObject("Excel.Application")
'    End If
'
'    Err.Clear
'
'    xlApp.Workbooks.Open (App.Path & "\AHD0101C.xls")
'
'    Set xlSheet = xlApp.Worksheets("Sheet1")
'    xlApp.Sheets("Sheet1").Select
'
'
'    xlApp.Range("O5").Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
'    xlApp.Range("C1").Value = text_cur_inv.Text + "�������ƽ�ⱨ��"
'    xlApp.Range("A5").Value = Mid(dtp_yy_mm.RawData, 1, 4) + "��" + _
'                              Mid(dtp_yy_mm.RawData, 5, 2) + "��" + _
'                              Mid(dtp_yy_mm.RawData, 7, 2) + "��"
'
'    Clipboard.Clear
'    ss1.Row = 1: ss1.Col = 2: ss1.Row2 = ss1.MaxRows: ss1.Col2 = 18
'    Clipboard.SetText ss1.Clip
'    xlApp.Range("A7").Select
'    xlApp.ActiveSheet.Paste
'
'
'    Clipboard.Clear
'    sRow = "A" & ss1.MaxRows + 6 & ":B" & ss1.MaxRows + 6
'    xlApp.Range(sRow).MERGECELLS = True
'    sRow = "A" & ss1.MaxRows + 6
'    xlApp.Range(sRow).Value = "��   ��:"
'    xlApp.Range(sRow).Font.Size = 12
'    xlApp.Range(sRow).Font.Bold = True
'    xlApp.ActiveSheet.Paste
'
'
'    ss1.ClearSelection
'    With xlApp.Application.FindFormat.Borders
'        .LineStyle = 1
'    End With
'
'    sRow = "A7:Q" & ss1.MaxRows + 6
'    xlApp.Range(sRow).Select
'    With xlApp.Selection.Borders
'        .LineStyle = 1
'    End With
''    xlApp.Columns("C:E").AutoFit
''    xlApp.Columns("J").AutoFit
'    Screen.MousePointer = vbDefault
'
'
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'    xlApp.DisplayAlerts = False
'    xlSheet.Close
'
'    xlApp.Quit
'
'    Set xlSheet = Nothing
'    Set xlApp = Nothing
'
'    Exit Sub
'
'ErrHandle:
'    MsgBox Error
'    Set xlSheet = Nothing
'    Set xlApp = Nothing
'    Screen.MousePointer = vbDefault
'End Sub


Private Sub subButtonHide()

    MDIMain.MenuTool.Buttons(4).Enabled = False    'Save
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    
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

    If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub



Private Sub Option1_Click()
    Text1.Text = "1"
End Sub

Private Sub Option2_Click()
    Text1.Text = "2"
End Sub
