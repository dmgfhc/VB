VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACA1130C 
   Caption         =   "合同兑现分析综合报表_ACA1130C"
   ClientHeight    =   10530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16440
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   16440
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   3615
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   15330
      _Version        =   393216
      _ExtentX        =   27040
      _ExtentY        =   6376
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
      MaxCols         =   17
      MaxRows         =   10
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "ACA1130C.frx":0000
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1720
      _Version        =   196609
      Begin VB.TextBox txt_shape 
         Alignment       =   2  'Center
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
         Left            =   13800
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "ss1"
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txt_del_date 
         Height          =   270
         Left            =   12360
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_marketing_cd 
         Height          =   270
         Left            =   13560
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_plt 
         Height          =   270
         Left            =   12360
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   360
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Caption         =   "用户交货月份"
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
      Begin InDate.UDate txt_del_to_date 
         Height          =   315
         Left            =   3060
         TabIndex        =   1
         Tag             =   "交货期"
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Text            =   "____-__"
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
         Mask            =   "%%%%-%%"
         MaxLength       =   7
      End
      Begin InDate.UDate txt_del_fr_date 
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Tag             =   "交货期"
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Text            =   "____-__"
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
         Mask            =   "%%%%-%%"
         MaxLength       =   7
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   360
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Caption         =   "报表编制日期"
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
      Begin InDate.UDate txt_report_date 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Tag             =   "报表编制日期"
         Top             =   600
         Width           =   1560
         _ExtentX        =   2752
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
      Begin Threed.SSCommand cmd_ord 
         Height          =   345
         Left            =   8520
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "合同兑现分析综合统计"
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4440
         Top             =   600
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "报表编制日期："
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
      Begin InDate.ULabel text_report_date 
         Height          =   315
         Left            =   5880
         Top             =   600
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
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
      Begin VB.Label Lab3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "汇总 导出"
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
         Index           =   0
         Left            =   8640
         TabIndex        =   10
         Top             =   550
         Width           =   1035
      End
      Begin VB.Label Lab3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "明细 导出"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   1
         Left            =   10080
         TabIndex        =   9
         Top             =   550
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "      **按月统计分析        起始月份与结束月份必须一致"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   90
      End
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   4620
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4560
      Width           =   15330
      _Version        =   393216
      _ExtentX        =   27040
      _ExtentY        =   8149
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
      MaxCols         =   25
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "ACA1130C.frx":0C82
   End
End
Attribute VB_Name = "ACA1130C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACA1130C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Cao Lei
'-- Coder             Cao Lei
'-- Date              2013.06.26
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  -------------------------------------------------------------------------------


Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public DEL_DATE As String           'Transfer to ACA1130C
Public PLT As String                'Transfer to ACA1130C
Public MARKETING_CD As String       'Transfer to ACA1130C

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection


Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2



Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
 '   FormType = "Msheet"
    FormType = "Refer"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_del_fr_date, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_del_to_date, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_report_date, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
          
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    
    
'    先注册1  查询明细使用
'    Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_del_date, "p", "n ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
           Call Gp_Ms_Collection(txt_plt, "p", "n ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
  Call Gp_Ms_Collection(txt_marketing_cd, "p", "n ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
       
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
       
       
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACA1130C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"
    
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, True)
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACA1130C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub


Private Sub cmd_ord_Click()
On Error GoTo cmd_ord_Error

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
'    Dim SMESG  As String
   
    Dim adoCmd As ADODB.Command
    
'    If TXT_DEL_FR_DATE.Text = "" Or TXT_DEL_TO_DATE.Text = "" Or TXT_DEL_FR_DATE.Text <> TXT_DEL_TO_DATE.Text Then
'       SMESG = "起始交货月份与结束交货月份必须一致！"
'       Call Gp_MsgBoxDisplay(SMESG)
'       Exit Sub
'    End If
        
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
              
    sQuery = "{call ACA1130P('" & (txt_del_fr_date.RawData) + "','" + Trim(txt_del_to_date.RawData) & "',?)}"
      
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'OS Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
   
        Call MsgBox(cmd_ord.Caption + "完成！", vbInformation, "系统提示信息")
        
        txt_report_date.Text = Date
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

cmd_ord_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("合同兑现分析统计错误: " & Error)

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
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("sc")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("sc")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("sc"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_ColGet(Proc_Sc("sc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "C-System.INI", Me.Name)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       cmd_ord.Visible = True
    Else
       cmd_ord.Visible = False
    End If
    
    txt_del_fr_date.Text = DateAdd("m", -1, Date)
    txt_del_to_date.Text = DateAdd("m", -1, Date)
    txt_report_date.Text = ""

    Screen.MousePointer = vbDefault

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "C-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
            
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

'
'Public Sub Spread_Can()
'
'    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
'
'End Sub


Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gf_Sp_Cls(Proc_Sc("SC2"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

    txt_del_fr_date.Text = Date
    txt_del_to_date.Text = Date
    txt_report_date.Text = ""
    text_report_date.Caption = ""
    
End Sub


Public Sub Form_Ref()

ss1.ReDraw = False

On Error GoTo Refer_Err

    Dim sMesg  As String
    Dim sQuery As String
    Dim iRow, i As Integer
    
    
    If txt_del_fr_date.Text <> txt_del_to_date.Text Then
       sMesg = "起始交货月份与结束交货月份必须一致！"
       Call Gp_MsgBoxDisplay(sMesg)
       Exit Sub
    End If
    
    '报表编制日期
    sQuery = "select max(t.rpt_date || t.rpt_time)  from CP_ORD_COMP_RPT t where  t.ord_del_date like  '" + Mid(Trim(txt_del_fr_date.RawData), 1, 6) + "%'"
    text_report_date.Caption = Gf_CodeFind(M_CN1, sQuery)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            ss1.OperationMode = OperationModeNormal
        End If
        
        ss1.ColWidth(0) = 15
        
         ss1.ROW = 1:   ss1.Col = 0:      ss1.Text = "中板研销"
         ss1.ROW = 2:   ss1.Col = 0:      ss1.Text = "中板出口"
         ss1.ROW = 3:   ss1.Col = 0:      ss1.Text = "中板汇总"
         ss1.ROW = 4:   ss1.Col = 0:      ss1.Text = "宽厚板研销"
         ss1.ROW = 5:   ss1.Col = 0:      ss1.Text = "宽厚板出口"
         ss1.ROW = 6:   ss1.Col = 0:      ss1.Text = "宽厚板汇总"
         ss1.ROW = 7:   ss1.Col = 0:      ss1.Text = "板卷研销"
         ss1.ROW = 8:   ss1.Col = 0:      ss1.Text = "板卷出口"
         ss1.ROW = 9:   ss1.Col = 0:      ss1.Text = "板卷汇总"
        ss1.ROW = 10:   ss1.Col = 0:      ss1.Text = "总计"
        
        For iRow = 1 To ss1.MaxRows
      ss1.ROW = iRow
      ss1.Col = 0
     If ss1.Text <> "" Then
       If ss1.Text = "中板汇总" Then
          For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
           Next
       End If
       
       If ss1.Text = "宽厚板汇总" Then
       For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
          Next
       End If
       
       If ss1.Text = "板卷汇总" Then
       For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
          Next
       End If
       
       If ss1.Text = "总计" Then
       For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HC0C0FF
          Next
       End If

      End If
        
    Next iRow
    
    
    Exit Sub
    
   
Refer_Err:
 
End Sub

'
'Public Sub Form_Pro()
'
'    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'
'End Sub

'
'Public Sub Form_Ins()
'
'    Call Gp_Sp_Ins(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
'
'End Sub


Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

'
'Public Sub Spread_Pst()
'
'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
'
'End Sub


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

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc2")("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
'   Call ss1_row_Click(Col, Row)

End Sub


Public Sub Form_Exc()
    
    If txt_shape.Text = "ss1" Then
     Call Gp_ACA1130C_Excel_D(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf txt_shape.Text = "ss2" Then
             Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
    

End Sub


Public Sub Form_Exit()
    Unload Me
End Sub


Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub


Private Sub Lab3_Click(Index As Integer)

    If Index = 0 Then
       txt_shape.Text = "ss1"
       Lab3(0).Caption = "汇总 导出"
       Lab3(1).Caption = "明细 导出"
       Lab3(0).BackColor = &HC0C0FF
       Lab3(1).BackColor = &HE0E0E0
    ElseIf Index = 1 Then
       txt_shape.Text = "ss2"
       Lab3(1).Caption = "明细 导出"
       Lab3(0).Caption = "汇总 导出"
       Lab3(1).BackColor = &HC0C0FF
       Lab3(0).BackColor = &HE0E0E0
    End If

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub


Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


'Private Sub ss2_LostFocus()
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub


Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

   Dim iRow As Long
   Dim iPlt As String
   Dim iMarket_cd As String
   Dim iDel_date As String
   
   iRow = ROW
   
   Call Gf_Sp_Cls(Proc_Sc("Sc2"))
   
   '清空ss2查询条件
   iPlt = ""
   iMarket_cd = ""
   iDel_date = ""
   
   '工厂
   ss1.ROW = iRow:  ss1.Col = 1:  iPlt = ss1.Text
   txt_plt.Text = iPlt
   
   '销售组别
   ss1.ROW = iRow:  ss1.Col = 2:  iMarket_cd = ss1.Text
   txt_marketing_cd.Text = iMarket_cd
   
   '交货月
   ss1.ROW = iRow:  ss1.Col = 3:  iDel_date = ss1.Text
   txt_del_date.Text = iDel_date
   
   
   If iPlt = "" Or iMarket_cd = "" Or iDel_date = "" Then
      Exit Sub
   End If
   
   If iRow > 0 Then
      
      If Gf_Sp_Refer(M_CN1, sc2, Mc2) Then
            ss2.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
      End If
   End If
   
   With ss2
   
        If .MaxRows < 1 Then
                Exit Sub
            End If
                 
        '未处理天数>6天 警示 红色
        For iRow = 1 To ss2.MaxRows
            .ROW = iRow
            .Col = 5
         If .Text > 6 Then
              Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, &HFF&)
          End If
            
        Next iRow
        
    End With
    
End Sub


Private Sub Gp_ACA1130C_Excel_D(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    'Call Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    With ss1
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then     '导出的表格从第几列数据开始
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .ROW = bLkrow1: .Row2 = bLkrow2
        
        Clipboard.SetText .Clip
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        xlSheet.Range("B1").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        sExlRange1 = ""
        
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .ROW = 1
            
            iExlCol = ColIndex
            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1") And _
               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
                If ColIndex > 104 Then
                    sExlRange1 = "E" & sExlRange1
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "D" & sExlRange1
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "C" & sExlRange1
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "B"
                    iExlCol = ColIndex - 26
                End If
                
            End If
        Next
        
    End With
    
   xlSheet.Range("A3").Value = "中板研销"
   xlSheet.Range("A4").Value = "中板出口"
   xlSheet.Range("A5").Value = "中板汇总"
   xlSheet.Range("A6").Value = "宽厚板研销"
   xlSheet.Range("A7").Value = "宽厚板出口"
   xlSheet.Range("A8").Value = "宽厚板汇总"
   xlSheet.Range("A9").Value = "板卷研销"
  xlSheet.Range("A10").Value = "板卷出口"
  xlSheet.Range("A11").Value = "板卷汇总"
  xlSheet.Range("A12").Value = "总计"
  
 xlSheet.Range("A1:R11 ").Borders.LineStyle = 1 '   增加边框


    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel", "W")

End Sub
