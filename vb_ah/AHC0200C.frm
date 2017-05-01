VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AHC0200C 
   Caption         =   "振华发货明细_AHC0200C "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_CUST 
      Height          =   270
      Left            =   5985
      MaxLength       =   6
      TabIndex        =   12
      Top             =   1155
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   150
      TabIndex        =   9
      Top             =   105
      Width           =   5310
      Begin VB.OptionButton Opt4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "上海振华港口机械（集团）股份有限公司"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   14
         Top             =   1245
         Width           =   3975
      End
      Begin VB.OptionButton Opt3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "上海金沿达"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   13
         Top             =   915
         Value           =   -1  'True
         Width           =   4020
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "上海金沿达钢材销售有限公司(振华港机）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   11
         Top             =   210
         Width           =   3990
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "上海致信钢材销售有限公司"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   10
         Top             =   555
         Width           =   3975
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   30
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "查询对象"
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
   Begin VB.TextBox TXT_MAXSEQ 
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   990
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TXT_TABLE2 
      Height          =   315
      Left            =   7425
      TabIndex        =   7
      Top             =   690
      Width           =   1215
   End
   Begin VB.TextBox TXT_TABLE 
      Height          =   270
      Left            =   10980
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "按订单号导出发运清单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   12720
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.CommandButton Print 
      Caption         =   "按合同号导出发运清单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11040
      TabIndex        =   4
      Top             =   240
      Width           =   2160
   End
   Begin VB.ComboBox Cbo_FH 
      Height          =   300
      Left            =   9765
      TabIndex        =   1
      Top             =   270
      Visible         =   0   'False
      Width           =   1875
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   6045
      Top             =   210
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "发货日期"
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
   Begin InDate.ULabel ULabel8 
      Height          =   300
      Left            =   9705
      Top             =   675
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Caption         =   "至"
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
      Height          =   7110
      Left            =   165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1875
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   12541
      _StockProps     =   64
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
      MaxCols         =   15
      MaxRows         =   20
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AHC0200C.frx":0000
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   8325
      Top             =   270
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "提单号"
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
   Begin InDate.UDate ud_MAX_DATE 
      Height          =   315
      Left            =   7425
      TabIndex        =   2
      Tag             =   "日期"
      Top             =   210
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.UDate ud_MIN_DATE 
      Height          =   315
      Left            =   8295
      TabIndex        =   3
      Tag             =   "日期"
      Top             =   675
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   6045
      Top             =   705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "再次打印表名"
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
Attribute VB_Name = "AHC0200C"
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
'-- Program ID        AHC0020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHANGLIN
'-- Coder             ZHANGLIN
'-- Date              2005.10.23
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

'    Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
'      Call Gp_Ms_Collection(ud_MIN_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(ud_MAX_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_CUST, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AHC0200C.P_REFER1", Key:="P-R"
    Sc1.Add Item:="AHC0200C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 13, True)
    Call Gp_Sp_ColHidden(ss1, 14, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub


Private Sub Command1_Click()

Dim fs, a, b, c
Dim txt_head, txt_tail, T_FILE_NAME As String

Dim MaxRows, iRowCount As Integer
      'Create file c:\PB\FP\JSKP.TXT
     Set fs = CreateObject("Scripting.FileSystemObject")                    '创建文件对象
     If fs.FOLDEREXISTS("D:\发运清单") Then                                    '创建目录
'        '
'     Else
'      If fs.FOLDEREXISTS("C:\PB") Then
'        Set b = fs.CreateFolder("C:\PB\FP")
'        Else
'        Set b = fs.CreateFolder("C:\PB")
'        Set b = fs.CreateFolder("C:\PB\FP")
'      End If
     End If
      
'       If fs.FileExists("c:\PB\FP\jskp.txt") Then
'       Else

'      T_FILE_NAME = Trim("C:\发运清单\" & ud_MAX_DATE.Text & "发运清单" & ".TXT")
    
      Set a = fs.CreateTextFile("D:\发运清单\" & ud_MAX_DATE.Text & "发运清单（订单号）" & ".TXT", True)

'      Set a = fs.CreateTextFile("C:\发运清单\发运清单.txt", True)
'       End If
'
      ' Set a = fs.opentextfile("c:\PB\FP\jskp.txt", 8, True, 0)

   '  Write to c:\PB\FP\JSKP.TXT
     If ss1.MaxRows = 0 Then
        Exit Sub
    
     End If
     
     For iRowCount = 1 To ss1.MaxRows
         txt_head = ""
         ss1.Row = iRowCount
         ss1.Col = 1
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Text + Chr$(9)
         
         ss1.Col = 3
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Text + Chr$(9)
         
         ss1.Col = 4
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Text + Chr$(9)

         ss1.Col = 5
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Value + Chr$(9)
         
         ss1.Col = 6
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Value + Chr$(9)

         ss1.Col = 9
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Value + Chr$(9)
         
         ss1.Col = 10
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Value + Chr$(9)

         ss1.Col = 7
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Text + Chr$(9)
         
         ss1.Col = 8
         If ss1.Text = "" Then ss1.Text = " "
         txt_head = txt_head + ss1.Text + Chr$(9)
        
'         ss1.Col = 11
'         If ss1.Text = "" Then ss1.Text = ""
'         txt_head = txt_head + ss1.Text + Chr$(9)
        
        a.WriteLine (txt_head)   '写第一行
                                 '写第二行
     Next iRowCount
     
 
    a.Close         '关闭文件
''    a.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
''    a.DisplayAlerts = True
''    xlSheet.Close

Call Gp_MsgBoxDisplay("导出完毕，文件保存在 D:\发运清单 ", "I")

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
''    Call subButtonHide
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    
''    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 1)
    
    Screen.MousePointer = vbDefault
    
    Opt1.Value = True
    Opt2.Value = False
    TXT_CUST.Text = "SH0030"

'    TXT_EMP = sUserID
    TXT_TABLE2.Text = ""
    TXT_MAXSEQ.Text = ""
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
    
''    Set iSumCol = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
''    Set iSumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
''        Call subButtonHide
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
'  '      rControl(1).SetFocus
    End If
 
    If Opt1 Then
       Call Opt1_Click
    Else
       Call Opt2_Click
    End If
    
    TXT_TABLE2.Text = ""
    TXT_TABLE.Text = ""
    TXT_MAXSEQ.Text = ""
 End Sub

Public Sub Form_Ref()

    Dim AdoRs As adodb.Recordset
    Dim sQuery      As String
    Dim i           As Integer

'
''On Error GoTo Refer_Err
''
''    Dim sMesg As String
''    Dim sQuery As String
''    sQuery = "{ CALL " + "AHD0110C.P_REFER" + "("
''    sQuery = sQuery + " '" + dtp_yy_mm.RawData + "'"
''    sQuery = sQuery + ")"
''    sQuery = sQuery + "}"
''
''
''
''    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
''
''    If dtp_yy_mm.RawData = "" Then
''       Call Gp_MsgBoxDisplay("请输入日期", "I")
''       Exit Sub
''    End If
''
''
'''    If Gf_Sp_Display(M_CN1, ss1, sQuery) Then
''    If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, 1, 4, iSumCnt, iSumCol, False) Then
''        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
''        Call subButtonHide
'''        Call Sp_AutoInsertSum
'''        Call Sp_AutoInsertSumGroup
''    End If
''
''
''    Exit Sub
''
''Refer_Err:

    Set AdoRs = New adodb.Recordset
       
    sQuery = "SELECT MAX(TABLE_NM) "
'    sQuery = sQuery & "'" & txt_ord_no_s & "',"
'    sQuery = sQuery & "'" & txt_ord_item_s & "',"
'    sQuery = sQuery & "'" & txt_PROD_CD.Text & "',"
'    sQuery = sQuery & "'" & txt_cur_inv_s & "') "
    sQuery = sQuery & " FROM hp_zh_shp "
    sQuery = sQuery & " WHERE SUBSTR(TABLE_NM,1,8) = '" & ud_MAX_DATE.RawData & "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        If AdoRs.Fields(0) <> "" Then
            TXT_TABLE.Text = Val(AdoRs.Fields(0)) + 1
        Else
            TXT_TABLE.Text = ud_MAX_DATE.RawData + "01"
        End If
    End If
    AdoRs.Close
    Set AdoRs = Nothing



On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Exit Sub
    End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
     TXT_TABLE2.Text = ""
     TXT_MAXSEQ.Text = ""
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    
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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, ss1.MaxCols - 1, lBlkrow1, lBlkrow1)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub Print_Click()
Dim AdoRs As adodb.Recordset
Dim sQuery      As String
Dim TAB_NM, TAB_SEQ As String
Dim COUNT As Integer


Dim fs, a, b, c
Dim txt_head, txt_head1, txt_tail, T_FILE_NAME As String

Dim MaxRows, iRowCount, jRowCount As Integer
      'Create file c:\PB\FP\JSKP.TXT
     Set fs = CreateObject("Scripting.FileSystemObject")                    '创建文件对象
     If fs.FOLDEREXISTS("D:\发运清单") Then                                    '创建目录
     Else
        Set b = fs.CreateFolder("D:\发运清单")
'      If fs.FOLDEREXISTS("C:\PB") Then
'        Set b = fs.CreateFolder("C:\PB\FP")
'        Else
'        Set b = fs.CreateFolder("C:\PB")
'        Set b = fs.CreateFolder("C:\PB\FP")
'      End If
     End If
      
''       If fs.FileExists("c:\PB\FP\jskp.txt") Then
''       Else
'
''      T_FILE_NAME = Trim("C:\发运清单\" & ud_MAX_DATE.Text & "发运清单" & ".TXT")
'
''      Set a = fs.CreateTextFile("D:\发运清单\" & ud_MAX_DATE.RawData & "_" & TXT_SEQ.Text & "发运清单（合同号）" & ".TXT", True)
'      If TXT_TABLE2.Text = "" Then
'        Set a = fs.CreateTextFile("D:\发运清单\" & TXT_TABLE.Text & "发运清单（合同号）" & ".TXT", True)
'      Else
'        Set a = fs.CreateTextFile("D:\发运清单\" & TXT_TABLE2.Text & "发运清单（合同号）" & ".TXT", True)
'      End If
''      Set a = fs.CreateTextFile("C:\发运清单\发运清单.txt", True)
''       End If
''
'      ' Set a = fs.opentextfile("c:\PB\FP\jskp.txt", 8, True, 0)
'
'   '  Write to c:\PB\FP\JSKP.TXT
     If ss1.MaxRows = 0 Then
        Exit Sub
    
     End If
     
 If TXT_TABLE2.Text = "" Then
     COUNT = 0
     
     For iRowCount = 1 To ss1.MaxRows
         txt_head = ""
         ss1.Row = iRowCount
         ss1.Col = 2
         If ss1.Text = "" Then
         
         Else
         
             ss1.Col = 11
             
             If ss1.Text <> "" Then
             
             Else
                 
                 ss1.Col = 0
                 ss1.Text = "Update"
                 COUNT = COUNT + 1
                 
                 ss1.Col = 11
                 ss1.Text = TXT_TABLE.Text
                 
                 ss1.Col = 12
                 ss1.Text = 1
                 
                 
                 ss1.Col = 13
                 ss1.Text = sUserID
                 
                 ss1.Col = 15
                 ss1.Text = ud_MAX_DATE.RawData
                 
                 ss1.Col = 1
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Text + Chr$(9)
                 
                 ss1.Col = 2
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Text + Chr$(9)
                 
                 ss1.Col = 4
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Value + Chr$(9)
                 
                 ss1.Col = 5
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Value + Chr$(9)
        
                 ss1.Col = 6
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Value + Chr$(9)
        
                 ss1.Col = 9
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Value + Chr$(9)
                 
                 ss1.Col = 10
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Value + Chr$(9)
        
                 ss1.Col = 7
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Text + Chr$(9)
                 
                 ss1.Col = 8
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head = txt_head + ss1.Text + Chr$(9)
                
'                 SS1.Col = 11
'                 If SS1.Text = "" Then SS1.Text = ""
'                 txt_head = txt_head + SS1.Text + Chr$(9)
                If COUNT = 1 Then
                   Set a = fs.CreateTextFile("D:\发运清单\" & TXT_TABLE.Text & "发运清单（合同号）" & ".TXT", True)
                End If
                a.WriteLine (txt_head)   '写第一行
                                         '写第二行
            End If
        End If
                                 
     Next iRowCount
     
     If COUNT > 0 Then
        Call Gp_MsgBoxDisplay("导出完毕，文件保存在 D:\发运清单 ", "I")
        a.Close  '关闭文件
    Else
        Call Gp_MsgBoxDisplay("没有可导出的数据", "I")
    End If
    
Else

    Set AdoRs = New adodb.Recordset
       
    sQuery = "SELECT MAX(SEQ) FROM HP_ZH_SHP WHERE TABLE_NM = '" & TXT_TABLE2.Text & "'"
'    sQuery = sQuery & "'" & txt_ord_item_s & "',"
'    sQuery = sQuery & "'" & txt_prod_cd.Text & "',"
'    sQuery = sQuery & "'" & txt_cur_inv_s & "') "
'    sQuery = sQuery & "FROM DUAL"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF And Not AdoRs.EOF Then
'    If AdoRs.RecordCount > 0 Then
        TXT_MAXSEQ.Text = Val(AdoRs.Fields(0) & "")
    End If
    AdoRs.Close
    Set AdoRs = Nothing
    
    COUNT = 0
     
     For jRowCount = 1 To ss1.MaxRows
         txt_head1 = ""
         ss1.Row = jRowCount
         ss1.Col = 11
         TAB_NM = ss1.Text
         ss1.Col = 12
         TAB_SEQ = ss1.Text
         If TXT_TABLE2.Text = TAB_NM And TXT_MAXSEQ.Text = TAB_SEQ Then
            ss1.Col = 0
            ss1.Text = "Update"
            COUNT = COUNT + 1
            
                 ss1.Col = 11
                 ss1.Text = TXT_TABLE2.Text
                 
                 ss1.Col = 12
                 ss1.Text = ss1.Text + 1
                 
                 
                 ss1.Col = 13
                 ss1.Text = sUserID
                 
'                 ss1.Col = 15
'                 ss1.Text = ud_MAX_DATE.RawData
                 
                 ss1.Col = 1
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Text + Chr$(9)
                 
                 ss1.Col = 2
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Text + Chr$(9)
                 
                 ss1.Col = 4
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Value + Chr$(9)
                 
                 ss1.Col = 5
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Value + Chr$(9)
        
                 ss1.Col = 6
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Value + Chr$(9)
        
                 ss1.Col = 9
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Value + Chr$(9)
                 
                 ss1.Col = 10
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Value + Chr$(9)
        
                 ss1.Col = 7
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Text + Chr$(9)
                 
                 ss1.Col = 8
                 If ss1.Text = "" Then ss1.Text = " "
                 txt_head1 = txt_head1 + ss1.Text + Chr$(9)
                
'                 SS1.Col = 11
'                 If SS1.Text = "" Then SS1.Text = ""
'                 txt_head = txt_head + SS1.Text + Chr$(9)
'                Set a = fs.CreateTextFile("D:\发运清单\" & TXT_TABLE.Text & "发运清单（合同号）" & ".TXT", True)
                If COUNT = 1 Then
                   Set a = fs.CreateTextFile("D:\发运清单\" & TXT_TABLE.Text & "发运清单（合同号）" & ".TXT", True)
                End If
                a.WriteLine (txt_head1)   '写第一行
                                         '写第二行
        End If
         
     Next jRowCount
     
     If COUNT > 0 Then
        Call Gp_MsgBoxDisplay("导出完毕，文件保存在 D:\发运清单 ", "I")
        a.Close  '关闭文件
    Else
        Call Gp_MsgBoxDisplay("没有可导出的数据", "I")
    End If

End If
     
'Call Gp_MsgBoxDisplay("导出完毕，文件保存在 D:\发运清单 ", "I")
TXT_TABLE2.Text = ""
TXT_MAXSEQ.Text = ""
Call Form_Pro
Call Form_Ref


End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
ss1.Row = ss1.ActiveRow
ss1.Col = 11
TXT_TABLE2.Text = ss1.Text


End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub
Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
'Private Sub AddOrdItem(Conn As ADODB.Connection)
'Dim sQuery As String
'
'    Dim sEdate1 As String
'    Dim sEdate2 As String
'
'    sEdate1 = Mid(ud_MIN_DATE.Text, 1, 4) + Mid(ud_MIN_DATE.Text, 6, 2) + Mid(ud_MIN_DATE.Text, 9, 2)
'    sEdate2 = Mid(ud_MAX_DATE.Text, 1, 4) + Mid(ud_MAX_DATE.Text, 6, 2) + Mid(ud_MAX_DATE.Text, 9, 2)
'If Trim(sEdate1) = "________" Then sEdate1 = "10000000"
'If Trim(sEdate2) = "________" Then sEdate2 = "99999999"
'
'
''    Screen.MousePointer = vbHourglass
'
'    sQuery = "Select distinct TRNS_NO from hp_shp_rslt "
'    sQuery = sQuery + " WHERE CUST_CD='SH0008'"
'    sQuery = sQuery + " AND   CAN_FL IS NULL"
'    sQuery = sQuery + " AND  SHP_date  between '" + sEdate1 + "' And '" + sEdate2 + "'"
'
'    Call Gf_ComboAdd(M_CN1, Cbo_FH, sQuery)
'
'End Sub
''Private Sub subButtonHide()
''
''    MDIMain.MenuTool.Buttons(4).Enabled = False    'Save
''    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
''    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
''    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
''    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
''
''    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
''    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
''
''End Sub

'Private Sub ud_MIN_DATE_LostFocus()
'  Call AddOrdItem(M_CN1)
'End Sub
'
'
'Private Sub ud_MAX_DATE_LostFocus()
'  Call AddOrdItem(M_CN1)
'End Sub


'Private Sub Print_Click()
'
'  Call create_txt_file(file_recordset())
'End Sub
'
'Function create_txt_file(file_recordset As Variant)
'Dim fs, a, b, c
'Dim txt_head, txt_tail, T_FILE_NAME As String
'
'Dim MaxRows, iRowCount As Integer
'      'Create file c:\PB\FP\JSKP.TXT
'     Set fs = CreateObject("Scripting.FileSystemObject")                    '创建文件对象
'     If fs.FOLDEREXISTS("D:\发运清单") Then                                    '创建目录
''        '
''     Else
''      If fs.FOLDEREXISTS("C:\PB") Then
''        Set b = fs.CreateFolder("C:\PB\FP")
''        Else
''        Set b = fs.CreateFolder("C:\PB")
''        Set b = fs.CreateFolder("C:\PB\FP")
''      End If
'     End If
'
''       If fs.FileExists("c:\PB\FP\jskp.txt") Then
''       Else
'    T_FILE_NAME = Trim("D:\发运清单\" & ud_MAX_DATE.Text & "发运清单" & ".TXT")
'
'      Set a = fs.CreateTextFile("T_FILE_NAME", True)
''       End If
''
'      ' Set a = fs.opentextfile("c:\PB\FP\jskp.txt", 8, True, 0)
'
'   '  Write to c:\PB\FP\JSKP.TXT
'     If IsEmpty(file_recordset) = True Then
'     MaxRows = 0
'       Else
'     MaxRows = UBound(file_recordset, 2) + 1
'     End If
'
'     For iRowCount = 0 To MaxRows - 1
'        txt_head = Trim(file_recordset(0, iRowCount))
'
'        a.WriteLine (txt_head)   '写第一行
'                                                                                            '写第二行
'     Next iRowCount
'
'
'    a.Close         '关闭文件
''    a.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
''    a.DisplayAlerts = True
''    xlSheet.Close
'
'End Function

'Function file_recordset()
'
'Dim sQuery As String
'Dim AdoRs As ADODB.Recordset
'Dim ArrayRecords_Head As Variant
'
'    Set AdoRs = New ADODB.Recordset
' '-----------------------------------------------------------------------------------------------------------------
'
'    sQuery = sQuery + "SELECT B.SHIP_ISP_NO||' '||A.ORD_NO||'   '||A.STLGRD||'  '"
'
'    sQuery = sQuery + "||A.THK ||'  '||A.WID||'   '"
'
'    sQuery = sQuery + "||SUBSTR(A.PROD_NO,1,10)||'   '||B.CERT_NO||'    '||' '"
'
'    sQuery = sQuery + "||COUNT(*)||'  '||A.WGT"
'
'    sQuery = sQuery + "   FROM hp_shp_rslt A ,qp_cert_head B  "
'    sQuery = sQuery + "   WHERE  A.TRNS_NO = B.TRNS_NO  AND  NVL(A.CAN_FL,'N') <>'Y' AND  A.CUST_CD = 'SH0005'"
'    sQuery = sQuery + "   AND A.TRNS_NO =  '" + Trim(Cbo_FH.Text) + "' "
'    sQuery = sQuery + "   AND A.SHP_DATE BETWEEN '" + Trim(ud_MIN_DATE.RawData) + "' AND '" + Trim(ud_MAX_DATE.RawData) + "'  "
'    sQuery = sQuery + "   GROUP BY SHIP_ISP_NO ,STLGRD,THK,WID, SUBSTR(A.PROD_NO,1,10),CERT_NO,WGT"
'
'   ' Ado Execute
'     AdoRs.Open sQuery, M_CN1, adOpenKeyset
'     If AdoRs.BOF Or AdoRs.EOF Then
'     Call Gp_MsgBoxDisplay("无相关记录", "I")
'
'     Exit Function
'     End If
'
'      ArrayRecords_Head = AdoRs.GetRows
'
'    '''-------------------------------------------------------------------------------------------------------------''
'
' file_recordset = ArrayRecords_Head
'End Function

Private Sub Opt1_Click()
        
    TXT_CUST.Text = "SH0030"
    
End Sub

Private Sub Opt2_Click()
    
    TXT_CUST.Text = "SH1170"
    
End Sub
        
Private Sub Opt3_Click()
        
    TXT_CUST.Text = "SH0028"
    
End Sub
        
Private Sub Opt4_Click()
        
    TXT_CUST.Text = "SH0008"
    
End Sub
