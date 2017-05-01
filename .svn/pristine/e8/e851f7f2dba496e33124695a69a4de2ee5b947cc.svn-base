VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AHQ0040C 
   Caption         =   "材质试验实绩确认_AHQ0040C"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_STDSPEC 
      Height          =   270
      Left            =   12360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_ORD_ITEM 
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
      Left            =   8190
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "发货指示号"
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox txt_prod_cd 
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
      Left            =   1065
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "产品"
      Top             =   90
      Width           =   570
   End
   Begin VB.TextBox txt_ORD_NO 
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
      Left            =   6780
      MaxLength       =   11
      TabIndex        =   2
      Top             =   90
      Width           =   1365
   End
   Begin VB.TextBox txt_PROD_NO 
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
      Left            =   3420
      MaxLength       =   14
      TabIndex        =   1
      Top             =   90
      Width           =   1725
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   2070
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "产品号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   5430
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "订单号"
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
      ForeColor       =   -2147483646
   End
   Begin FPSpread.vaSpread ss4 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   2143
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
      MaxCols         =   17
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AHQ0040C.frx":0000
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   120
      Top             =   90
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   "品种"
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5925
      Left            =   105
      TabIndex        =   6
      Top             =   3090
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   10451
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AHQ0040C.frx":07BD
      Begin FPSpread.vaSpread SS2 
         Height          =   5925
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   6240
         _Version        =   393216
         _ExtentX        =   11007
         _ExtentY        =   10451
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
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "AHQ0040C.frx":080F
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   5925
         Left            =   6330
         TabIndex        =   8
         Top             =   0
         Width           =   8700
         _Version        =   393216
         _ExtentX        =   15346
         _ExtentY        =   10451
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
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
         MaxCols         =   5
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AHQ0040C.frx":0B6B
      End
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   1245
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   2196
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
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
      MaxCols         =   15
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AHQ0040C.frx":0F2C
   End
End
Attribute VB_Name = "AHQ0040C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      材质试验实绩输入
'-- Program ID        AQC0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HAN.Y.S
'-- Coder             ZENG.W
'-- Date              2005.10. 25
'-- Description       材质试验实绩输入
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

Dim pColumn12 As New Collection      'Spread Primary Key Collection
Dim nColumn12 As New Collection      'Spread necessary Column Collection
Dim mColumn12 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn12 As New Collection      'Spread Insert Column Collection
Dim aColumn12 As New Collection      'Master -> Spread Column Collection
Dim lColumn12 As New Collection      'Spread Lock Column Collection

Dim pColumn13 As New Collection      'Spread Primary Key Collection
Dim nColumn13 As New Collection      'Spread necessary Column Collection
Dim mColumn13 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn13 As New Collection      'Spread Insert Column Collection
Dim aColumn13 As New Collection      'Master -> Spread Column Collection
Dim lColumn13 As New Collection      'Spread Lock Column Collection

Dim pColumn14 As New Collection      'Spread Primary Key Collection
Dim nColumn14 As New Collection      'Spread necessary Column Collection
Dim mColumn14 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn14 As New Collection      'Spread Insert Column Collection
Dim aColumn14 As New Collection      'Master -> Spread Column Collection
Dim lColumn14 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection
Dim Sc3 As New Collection
Dim Sc4 As New Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim arrChem(3, 51) As String
Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_PROD_CD, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_PROD_NO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ORD_NO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_ord_item, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AHQ0040C.P_REFER", Key:="P-R"
'    sc1.Add Item:="AQC0040C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
'     Call SS1.AddCellSpan(5, 0, 1, 2)

      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(SS2, 1, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(SS2, 2, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(SS2, 3, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(SS2, 4, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     
     'Spread_Collection
    Sc2.Add Item:=SS2, Key:="Spread"
    Sc2.Add Item:="AHQ0040C.P_SREFER_1", Key:="P-R"
    Sc2.Add Item:=pColumn12, Key:="pColumn"
    Sc2.Add Item:=nColumn12, Key:="nColumn"
    Sc2.Add Item:=aColumn12, Key:="aColumn"
    Sc2.Add Item:=mColumn12, Key:="mColumn"
    Sc2.Add Item:=iColumn12, Key:="iColumn"
    Sc2.Add Item:=lColumn12, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=SS2.MaxCols, Key:="Last"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     
     'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AHQ0040C.P_SREFER_2", Key:="P-R"
    Sc3.Add Item:=pColumn13, Key:="pColumn"
    Sc3.Add Item:=nColumn13, Key:="nColumn"
    Sc3.Add Item:=aColumn13, Key:="aColumn"
    Sc3.Add Item:=mColumn13, Key:="mColumn"
    Sc3.Add Item:=iColumn13, Key:="iColumn"
    Sc3.Add Item:=lColumn13, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 5, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 6, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 7, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 8, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     Call Gp_Sp_Collection(ss4, 9, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 10, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 11, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 12, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 13, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 14, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 15, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 16, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 17, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
     
     'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AHQ0040C.P_REFER2", Key:="P-R"
    Sc4.Add Item:=pColumn14, Key:="pColumn"
    Sc4.Add Item:=nColumn14, Key:="nColumn"
    Sc4.Add Item:=aColumn14, Key:="aColumn"
    Sc4.Add Item:=mColumn14, Key:="mColumn"
    Sc4.Add Item:=iColumn14, Key:="iColumn"
    Sc4.Add Item:=lColumn14, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss4, 1, True)
    Call Gp_Sp_ColHidden(ss4, 2, True)
    Call Gp_Sp_ColHidden(ss4, 12, True)
    Call Gp_Sp_ColHidden(ss4, 14, True)
    Call Gp_Sp_ColHidden(ss1, 2, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
'    Call Gp_Sp_BlockColor(SS1, 2, SS1.MaxCols, 1, SS1.MaxRows, , &HFFFF&)


End Sub

'Private Sub cmd_AllCheck_Click()
'    Dim i       As Integer
'    Dim sAllChk As String
'
'    If SS1.MaxRows < 1 Or SS1.Row = 0 Then Exit Sub
'
'    If cmd_AllCheck.Caption = "全部确认" Then
'        sAllChk = "ALL"
'    Else
'        sAllChk = ""
'    End If
'
'    If Gf_Sc_Authority(sAuthority, "U") Then
'
'        For i = 1 To SS1.MaxRows
'            SS1.Row = i
'            If sAllChk = "ALL" Then
'                SS1.Col = 1
'                SS1.Text = 1
'                SS1.Col = 0
'                SS1.Text = "Update"
'                cmd_AllCheck.Caption = "全部取消"
'            Else
'                SS1.Col = 1
'                SS1.Text = 0
'                SS1.Col = 0
'                SS1.Text = ""
'                cmd_AllCheck.Caption = "全部确认"
'            End If
'        Next i
'
'    End If
'
'End Sub

'Private Sub MenuToolSet()
'
'    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
'    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
'    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
'    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
'    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
'    MDIMain.MenuTool.Buttons(14).Enabled = False
'
'End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
'    Call MenuToolSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Track:
'    Dim oCodeName As Object
'    Dim sCode As String
'
'    Select Case Me.ActiveControl.Name
'
'        Case "txt_STDSPEC"              '标准
'            sCode = "STDSPEC"
'            Set oCodeName = txt_STDSPEC_NAME
'    End Select
'
'    If sCode = "" Then Exit Sub
'
'    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
'
'    Set oCodeName = Nothing
'
'Err_Track:
'
'    Set oCodeName = Nothing
'
'End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
'    Call MenuToolSet

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(ss1)
    Call Gp_Sp_Setting(SS2)
    Call Gp_Sp_Setting(ss3)
    Call Gp_Sp_Setting(ss4)
    Call Gp_Sp_ReadOnlySet(SS2)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(SS2, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss4, "Q-System.INI", Me.Name)
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Set iColumn12 = Nothing
    Set pColumn12 = Nothing
    Set lColumn12 = Nothing
    Set nColumn12 = Nothing
    Set mColumn12 = Nothing
    Set aColumn12 = Nothing
    
    Set iColumn13 = Nothing
    Set pColumn13 = Nothing
    Set lColumn13 = Nothing
    Set nColumn13 = Nothing
    Set mColumn13 = Nothing
    Set aColumn13 = Nothing
    
    Set iColumn14 = Nothing
    Set pColumn14 = Nothing
    Set lColumn14 = Nothing
    Set nColumn14 = Nothing
    Set mColumn14 = Nothing
    Set aColumn14 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set Sc4 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(Sc2)
        Call Gf_Sp_Cls(Sc3)
        Call Gf_Sp_Cls(Sc4)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If

End Sub

Public Sub Form_Ref()
    Dim iRow, iCol  As Integer
    Dim sQuery      As String
    Dim sMesg       As String
    Dim AdoRs       As ADODB.Recordset

    On Error GoTo Refer_Err
    

    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
       If Gf_Sp_Refer(M_CN1, Sc4, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            ss4.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       End If
'        Call MenuToolSet
    End If
    
    Call Gf_Sp_Cls(Sc2)
    Call Gf_Sp_Cls(Sc3)
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub
    
    Set AdoRs = New ADODB.Recordset
       
    sQuery = "SELECT Gf_AQC1711P( "
    sQuery = sQuery & "'" & txt_PROD_NO.Text & "',"
    sQuery = sQuery & "'" & txt_PROD_CD.Text & "',"
    sQuery = sQuery & "'" & txt_ORD_NO.Text & "',"
    sQuery = sQuery & "'" & txt_ord_item.Text & "') "
    sQuery = sQuery & "FROM DUAL"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF And Not AdoRs.EOF Then
'    If AdoRs.RecordCount > 0 Then

       MDIMain.StatusBar1.Panels(1) = "提示信息： " & AdoRs.Fields(0)
       
    End If
    AdoRs.Close
    Set AdoRs = Nothing

Refer_Err:
    
    Screen.MousePointer = vbDefault

End Sub

'Public Sub Form_Pro()
'
'    Call DataSave("1")
'
'End Sub

'Private Sub cmd_PIC_Click()
'
'   Call DataSave("2")
'
'End Sub

'Public Sub DataSave(SaveFL As String)
'    Dim iRow, iCol As Integer
'
'    sc1.Remove ("P-M")
'    If SaveFL = "1" Then
'        sc1.Add Item:="AQC0040C.P_MODIFY1", Key:="P-M"
'    Else
'        sc1.Add Item:="AQC0040C.P_MODIFY2", Key:="P-M"
'    End If
'
'    With SS1
'       For iRow = 1 To .MaxRows
'           .Row = iRow
'           .Col = 0
'           If .Text = "Update" Then
'              .Col = 7
'              .Text = sUserID
'           End If
'       Next iRow
'    End With
'
'    If Gf_Sp_Process(M_CN1, sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'
'    SS1.OperationMode = OperationModeNormal
''    Call MenuToolSet
'
'    If SS1.MaxRows < 1 Or SS1.ActiveRow = 0 Then Exit Sub
'
'    With SS1
'         For iRow = 1 To .MaxRows
'            .Row = iRow
'            .Col = 5
'            If .Text = "Y" Then
'               Call Gp_Sp_BlockColor(SS1, 2, SS1.MaxCols, iRow, iRow, , &HFFFF&)
'            Else
'                Call Gp_Sp_BlockColor(SS1, 2, SS1.MaxCols, iRow, iRow, , &H80000005)
'            End If
'         Next iRow
'    End With
'
'End Sub


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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
            
    Dim sQuery          As String
    Dim sMesg           As String
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant
    Dim arr             As Variant
    Dim SMP_NO, smp_loc As Variant
 
 On Error GoTo Error_Rtn
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    If ss1.MaxRows < 1 Or Row = 0 Or Col = 1 Then Exit Sub
    
'    If Col = 0 Then
'
'        Unload AQC0080C
'
'        SS1.Row = Row
'        SS1.Col = 8
'        AQC0080C.txt_ORD_NO = Trim(SS1.Text)
'        SS1.Col = 9
'        AQC0080C.TXT_ORD_ITEM = Trim(SS1.Text)
'
'        AQC0080C.Show
'        AQC0080C.Form_Ref
'
'        Exit Sub
'
'    End If
    With ss1
        .Col = 3
        .Row = .ActiveRow
        SMP_NO = .Text
        .Col = 3
        TXT_STDSPEC = .Text
'        .Col = 6
'        sdb_ORD_WID = .Text
'        .Col = 10
'        txt_TIME.RawData = .Text
    End With
    
    SS2.MaxRows = 0
    ss3.MaxRows = 0
    
    ss1.ReDraw = False
    SS2.ReDraw = False
    ss3.ReDraw = False
    ss4.ReDraw = False
    
    Set AdoRs = New ADODB.Recordset
    
    If Trim(SMP_NO) = "" Then
        sQuery = "{call AHQ0040C.P_SREFER_1('" + Trim(txt_PROD_NO.Text) + "','" & Trim(txt_ORD_NO.Text) & "','" & Trim(txt_ord_item.Text) & "')}"
    Else
        sQuery = "{call AHQ0040C.P_SREFER_1('" + Trim(SMP_NO) + "','" & Trim(txt_ORD_NO.Text) & "','" & Trim(txt_ord_item.Text) & "')}"
    End If
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView2(ArrayRecords)
        Erase ArrayRecords
    End If
    
'    Call Gp_Sp_EvenRowBackcolor(SS2)

    If Trim(SMP_NO) <> "" Then
        sQuery = "{call AHQ0040C.P_SREFER_2('" + Trim(SMP_NO) + "','" & Trim(txt_ORD_NO.Text) & "','" & Trim(txt_ord_item.Text) & "')}"
                        
        AdoRs.Close
        
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
            
        If Not (AdoRs.BOF And AdoRs.EOF) Then
            ArrayRecords = AdoRs.GetRows
            Call subSpreadView1(ArrayRecords)
            Erase ArrayRecords
        End If

        sQuery = "{call AHQ0040C.P_SREFER_3('" + Trim(SMP_NO) + "','" & Trim(txt_ORD_NO.Text) & "','" & Trim(txt_ord_item.Text) & "')}"
        
        AdoRs.Close
                        
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
            
        If Not (AdoRs.BOF And AdoRs.EOF) Then
            ArrayRecords = AdoRs.GetRows
            Call subSpreadView3(ArrayRecords)
            Erase ArrayRecords
        End If
    End If
    

'    Call Gp_Sp_EvenRowBackcolor(ss3)
    
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    ss1.ReDraw = True
    SS2.ReDraw = True
    ss3.ReDraw = True
    ss4.ReDraw = True
    
    Exit Sub
    
Error_Rtn:
    
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    Screen.MousePointer = vbDefault
    ss1.ReDraw = True
    SS2.ReDraw = True
    ss3.ReDraw = True
    ss4.ReDraw = True
End Sub
'Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
'
'    Dim sQuery          As String
'    Dim sMesg           As String
'    Dim AdoRs           As ADODB.Recordset
'    Dim ArrayRecords    As Variant
'    Dim arr             As Variant
'    Dim SMP_NO, smp_loc As Variant
'
' On Error GoTo Error_Rtn
'
'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
'
'    If SS1.MaxRows < 1 Or Row = 0 Or Col = 1 Then Exit Sub
'
'
'    With SS1
'        .Col = 1
'        .Row = .ActiveRow
'        txt_ORD_NO.Text = .Text
'        .Col = 2
'        txt_ORD_ITEM.Text = .Text
''        .Col = 4
''        TXT_STDSPEC = .Text
''        .Col = 6
''        sdb_ORD_WID = .Text
''        .Col = 10
''        txt_TIME.RawData = .Text
'    End With
'
'    SS2.MaxRows = 0
'    ss3.MaxRows = 0
'
'    SS1.ReDraw = False
'    SS2.ReDraw = False
'    ss3.ReDraw = False
'    ss4.ReDraw = False
'
'    Set AdoRs = New ADODB.Recordset
'    sQuery = "{call AHQ0040C.P_SREFER_1('" + Trim(SMP_NO) + "')}"
'
'    AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'    If Not (AdoRs.BOF And AdoRs.EOF) Then
'        ArrayRecords = AdoRs.GetRows
'        Call subSpreadView2(ArrayRecords)
'        Erase ArrayRecords
'    End If
'
''    Call Gp_Sp_EvenRowBackcolor(SS2)
'
'    sQuery = "{call AQC0040C.P_SREFER_2('" + Trim(SMP_NO) + "')}"
'
'    AdoRs.Close
'
'    AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'    If Not (AdoRs.BOF And AdoRs.EOF) Then
'        ArrayRecords = AdoRs.GetRows
'        Call subSpreadView1(ArrayRecords)
'        Erase ArrayRecords
'    End If
'
'    sQuery = "{call AQC0040C.P_SREFER_3('" + Trim(SMP_NO) + "')}"
'
'    AdoRs.Close
'
'    AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'    If Not (AdoRs.BOF And AdoRs.EOF) Then
'        ArrayRecords = AdoRs.GetRows
'        Call subSpreadView3(ArrayRecords)
'        Erase ArrayRecords
'    End If
'
'
''    Call Gp_Sp_EvenRowBackcolor(ss3)
'
'    Set AdoRs = Nothing
'    Set ArrayRecords = Nothing
'    SS1.ReDraw = True
'    SS2.ReDraw = True
'    ss3.ReDraw = True
'    ss4.ReDraw = True
'
'    Exit Sub
'
'Error_Rtn:
'
'    Set AdoRs = Nothing
'    Set ArrayRecords = Nothing
'    Screen.MousePointer = vbDefault
'    SS1.ReDraw = True
'    SS2.ReDraw = True
'    ss3.ReDraw = True
'    ss4.ReDraw = True
'End Sub
''Private Sub InputEditCheck()
''
''    If SS1.ActiveCol <> 1 Then
''        pControl(1).SetFocus
''    End If
''
''End Sub
''
''Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
''    Call InputEditCheck
''End Sub


'Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
'
'    Call InputEditCheck
'
'    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
'
'    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'    End If
'
'    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True
'
'End Sub

'Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    ss1.Row = ss1.ActiveRow + 1
'    Call ss1_Click(ss1.Col, ss1.ActiveRow + 1)
'End Sub

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
Private Sub subSpreadView1(ByVal strArr As Variant)

    Dim I           As Integer
    Dim iRow        As Integer
    Dim sMatr(165)   As String
    
    If UBound(strArr, 2) < 0 Then Exit Sub
        
    sMatr(0) = "屈服点实绩                            "
    sMatr(1) = "拉伸规定总伸长应力实绩                "
    sMatr(2) = "抗拉强度实绩                          "
    sMatr(3) = "屈强比实绩                            "
    sMatr(4) = "断后伸长率实绩                        "
    sMatr(5) = "断面收缩率实绩1                       "
    sMatr(6) = "断面收缩率实绩2                       "
    sMatr(7) = "断面收缩率实绩3                       "
    sMatr(8) = "断面收缩率实绩平均                    "
    sMatr(9) = "冷弯试验实绩                          "
    sMatr(10) = "冲击试验温度                         "
    sMatr(11) = "冲击试样尺寸                         "
    sMatr(12) = "冲击试验实绩 1                       "
    sMatr(13) = "冲击试验实绩 2                       "
    sMatr(14) = "冲击试验实绩 3                       "
    sMatr(15) = "冲击试验实绩 4                       "
    sMatr(16) = "冲击试验实绩 5                       "
    sMatr(17) = "冲击试验实绩 6                       "
    sMatr(18) = "冲击试验实绩平均                     "
   
    sMatr(19) = "冲击剪切面积实绩平均                 "
    sMatr(20) = "冲击剪切面积实绩 1                   "
    sMatr(21) = "冲击剪切面积实绩 2                   "
    sMatr(22) = "冲击剪切面积实绩 3                   "
    sMatr(23) = "冲击剪切面积实绩 4                   "
    sMatr(24) = "冲击剪切面积实绩 5                   "
    sMatr(25) = "冲击剪切面积实绩 6                   "
    sMatr(26) = "时效冲击试验温度                     "
    sMatr(27) = "时效冲击试样尺寸                     "
    sMatr(28) = "时效冲击功实绩1                      "
    sMatr(29) = "时效冲击功实绩2                      "
    sMatr(30) = "时效冲击功实绩3                      "
    sMatr(31) = "时效冲击功实绩4                      "
    sMatr(32) = "时效冲击功实绩5                      "
    sMatr(33) = "时效冲击功实绩6                      "
    sMatr(34) = "时效冲击实绩平均                     "
                  
    sMatr(35) = "时效冲击纤维断面率实绩               "
    sMatr(36) = "重力撕裂温度                         "
    sMatr(37) = "重力撕裂实绩1                        "
    sMatr(38) = "重力撕裂实绩2                        "
    sMatr(39) = "重力撕裂实绩平均                     "
    sMatr(40) = "硬度实绩                             "
    sMatr(41) = "拉伸规定非比例伸长应力实绩           "
    sMatr(42) = "拉伸规定残余伸长应力实绩实绩         "
    sMatr(43) = "高温拉伸屈服强度实绩                 "
    sMatr(44) = "高温拉伸抗拉强度实绩                 "
    sMatr(45) = "高温拉伸断面收缩率实绩1              "
'20090806 SUN BIN
    sMatr(46) = "高温拉伸断面收缩率实绩2              "
    sMatr(47) = "高温拉伸断面收缩率实绩3              "
    sMatr(48) = "高温拉伸断面收缩率实绩平均           "
'20090806 SUN BIN END
    sMatr(49) = "高温拉伸断后伸长率实绩               "
    sMatr(50) = "高温拉伸规定非比例伸长应力实绩       "
    sMatr(51) = "高温拉伸规定残余伸长应力实绩         "
    sMatr(52) = "焊接硬度实绩                         "
    sMatr(53) = "焊缝弯曲实绩                         "
    sMatr(54) = "反复弯曲实绩                         "
    sMatr(55) = "锻平试验实绩                         "
    sMatr(56) = "抗氢裂能力CSR实绩                    "
    sMatr(57) = "抗氢裂能力CLR实绩                    "
    sMatr(58) = "抗氢裂能力CWR实绩                    "
    sMatr(59) = "硫化物腐蚀裂纹实绩                   "
    sMatr(60) = "追加冲击试验温度                     "
    sMatr(61) = "追加击试样尺寸                       "
    sMatr(62) = "追加冲击试验实绩平均                 "
    sMatr(63) = "追加冲击试验实绩 1                   "
    sMatr(64) = "追加冲击试验实绩 2                   "
    sMatr(65) = "追加冲击试验实绩 3                   "
    sMatr(66) = "追加冲击试验实绩 4                   "
    sMatr(67) = "追加冲击试验实绩 5                   "
    sMatr(68) = "追加冲击试验实绩 6                   "
    sMatr(69) = "追加冲击剪切面积实绩平均             "
    sMatr(70) = "追加冲击剪切面积实绩 1               "
    sMatr(71) = "追加冲击剪切面积实绩 2               "
    sMatr(72) = "追加冲击剪切面积实绩 3               "
    sMatr(73) = "追加冲击剪切面积实绩 4               "
    sMatr(74) = "追加冲击剪切面积实绩 5               "
    sMatr(75) = "追加冲击剪切面积实绩 6               "
    sMatr(76) = "追加时效冲击试验温度                 "
    sMatr(77) = "追加时效冲击试样尺寸                 "
    sMatr(78) = "追加时效冲击实绩平均                 "
    sMatr(79) = "追加时效冲击功实绩1                  "
    sMatr(80) = "追加时效冲击功实绩2                  "
    sMatr(81) = "追加时效冲击功实绩3                  "
    sMatr(82) = "追加时效冲击功实绩4                  "
    sMatr(83) = "追加时效冲击功实绩5                  "
    sMatr(84) = "追加时效冲击功实绩6                  "
    sMatr(85) = "追加时效冲击纤维断面率实绩           "
    sMatr(86) = "晶粒度实绩                           "
    sMatr(87) = "脱碳层实绩                           "
    sMatr(88) = "硫印实绩                             "
    sMatr(89) = "断口检验实绩1                        "
    sMatr(90) = "断口检验实绩2                        "
    sMatr(91) = "断口检验实绩3                        "
    sMatr(92) = "断口检验实绩4                        "
    sMatr(93) = "断口检验实绩5                        "
    sMatr(94) = "酸浸检验实绩1                        "
    sMatr(95) = "酸浸检验实绩2                        "
    sMatr(96) = "酸浸检验实绩3                        "
    sMatr(97) = "酸浸检验实绩4                        "
    sMatr(98) = "酸浸检验实绩5                        "
    sMatr(99) = "带状组织实绩                         "
    sMatr(100) = "淬透性试验实绩1                     "
    sMatr(101) = "淬透性试验实绩2                      "
    sMatr(102) = "淬透性试验实绩3                      "
    sMatr(103) = "非金属夹杂物(粗)实绩1                "
    sMatr(104) = "非金属夹杂物(粗)实绩2                "
    sMatr(105) = "非金属夹杂物(粗)实绩3                "
    sMatr(106) = "非金属夹杂物(粗)实绩4                "
    sMatr(107) = "非金属夹杂物(细)实绩1                "
    sMatr(108) = "非金属夹杂物(细)实绩2                "
    sMatr(109) = "非金属夹杂物(细)实绩3                "
    sMatr(110) = "非金属夹杂物(细)实绩4                "
    sMatr(111) = "奥氏体晶粒度实绩                     "
    sMatr(112) = "DS类非金属夹杂实绩                   "
    sMatr(113) = "TIN类非金属夹杂实绩                  "
'20090804 sun bin start
    sMatr(114) = "追加屈服点实绩                           "
    sMatr(115) = "追加拉伸规定总伸长应力实绩               "
    sMatr(116) = "追加抗拉强度实绩                         "
    sMatr(117) = "追加屈强比实绩                           "
    sMatr(118) = "追加断后伸长率实绩                       "
    sMatr(119) = "追加断面收缩率实绩1                      "
    sMatr(120) = "追加断面收缩率实绩2                      "
    sMatr(121) = "追加断面收缩率实绩3                      "
    sMatr(122) = "追加断面收缩率实绩平均                   "
    sMatr(123) = "追加冷弯试验实绩                         "
    sMatr(124) = "追加硬度实绩                             "
    sMatr(125) = "追加拉伸规定非比例伸长应力实绩           "
    sMatr(126) = "追加拉伸规定残余伸长应力实绩实绩         "
    sMatr(127) = "追加高温拉伸屈服强度实绩                 "
    sMatr(128) = "追加高温拉伸抗拉强度实绩                 "
    sMatr(129) = "追加高温拉伸断面收缩率实绩1              "
'20090806 sun bin start
    sMatr(130) = "追加高温拉伸断面收缩率实绩2              "
    sMatr(131) = "追加高温拉伸断面收缩率实绩3              "
    sMatr(132) = "追加高温拉伸断面收缩率实绩平均           "
'20090806 sun bin end
    sMatr(133) = "追加高温拉伸断后伸长率实绩               "
    sMatr(134) = "追加高温拉伸规定非比例伸长应力实绩       "
    sMatr(135) = "追加高温拉伸规定残余伸长应力实绩         "
'20090804 sun bin end
' edit for lou by geng
    sMatr(136) = "厚度方向面缩率实绩1                     "
    sMatr(137) = "厚度方向面缩率实绩2                     "
    sMatr(138) = "厚度方向面缩率实绩3                     "
    sMatr(139) = "厚度方向面缩率均值实绩                  "
    sMatr(140) = "高温厚度方向面缩率实绩1                 "
    sMatr(141) = "高温厚度方向面缩率实绩2                 "
    sMatr(142) = "高温厚度方向面缩率实绩3                 "
    sMatr(143) = "高温厚度方向面缩率均值实绩              "
    sMatr(144) = "侧膨胀值均值实绩                       "
    sMatr(145) = "侧膨胀值实绩1                      "
    sMatr(146) = "侧膨胀值实绩2                      "
    sMatr(147) = "侧膨胀值实绩3                      "
    sMatr(148) = "侧膨胀值实绩4                      "
    sMatr(149) = "侧膨胀值实绩5                      "
    sMatr(150) = "侧膨胀值实绩6                      "
    sMatr(151) = "追加侧膨胀值均值实绩                   "
    sMatr(152) = "追加侧膨胀值实绩1                      "
    sMatr(153) = "追加侧膨胀值实绩2                      "
    sMatr(154) = "追加侧膨胀值实绩3                      "
    sMatr(155) = "追加侧膨胀值实绩4                      "
    sMatr(156) = "追加侧膨胀值实绩5                      "
    sMatr(157) = "追加侧膨胀值实绩6                      "
'20110217 GENGXUEYU for 抗大
    sMatr(158) = "均匀变形伸长率UEL                       "
    sMatr(159) = "追加均匀变形伸长率UEL                   "
    sMatr(160) = "追加应力比项目1                         "
    sMatr(161) = "追加应力比项目2                         "
    sMatr(162) = "追加应力比项目3                         "
    sMatr(163) = "追加应力比项目4                         "
    sMatr(164) = "追加应力比项目5                         "
    With ss3
        .MaxRows = 165
    
        For I = 1 To 165
            .Row = I
            .Col = 1: .Text = sMatr(I - 1)
        Next I
                
        For I = 1 To UBound(strArr, 1) + 1
        
            .Row = I: .Col = 4
            .Text = NullCheck(strArr(I - 1, 0), "")
            
        Next I
    End With

    With ss3
        .MaxRows = 165
    
        For I = 1 To 165
            .Row = I
            .Col = 1: .Text = sMatr(I - 1)
        Next I
                
        For I = 1 To UBound(strArr, 1) + 1
        
            .Row = I: .Col = 4
            .Text = NullCheck(strArr(I - 1, 0), "")
            
        Next I
    End With
  
End Sub

Private Sub subSpreadView3(ByVal strArr As Variant)

    Dim I                     As Integer
    Dim iRow                  As Integer
    Dim sMatr(3, 165)         As Variant
    Dim sMatrCON(6, 165)      As Variant
    Dim sMin, sMax, sFL, sRE  As Variant
    
    If UBound(strArr, 2) < 0 Then Exit Sub
      
    If UBound(strArr, 2) = 0 Then
        For I = 0 To 164
            sMatr(0, I) = NullCheck(strArr(I, 0), "")
        Next I
        
        For I = 0 To 164
            sMatr(1, I) = NullCheck(strArr(I + 165, 0))
        Next I
    
        For I = 0 To 164
            sMatr(2, I) = NullCheck(strArr(I + 330, 0))
        Next I
        
        
        With ss3
                
            For I = 1 To 165
                .Row = I
                .Col = 2: .Text = sMatr(1, I - 1)
                .Col = 3: .Text = sMatr(2, I - 1)
                .Col = 5: .Text = sMatr(0, I - 1)
            Next I
         End With
    End If
     
    If UBound(strArr, 2) = 1 Then
        For I = 0 To 164
            sMatrCON(0, I) = NullCheck(strArr(I, 0), "")
            sMatrCON(3, I) = NullCheck(strArr(I, 1), "")
        Next I
        
        For I = 0 To 164
            sMatrCON(1, I) = NullCheck(strArr(I + 165, 0))
            sMatrCON(4, I) = NullCheck(strArr(I + 165, 1))
        Next I
    
        For I = 0 To 164
            sMatrCON(2, I) = NullCheck(strArr(I + 330, 0))
            sMatrCON(5, I) = NullCheck(strArr(I + 330, 1))
        Next I
        
            
        For I = 1 To 165
            If sMatrCON(0, I - 1) = "A" Or sMatrCON(0, I - 1) = "B" Then
                If sMatrCON(3, I - 1) = "A" Or sMatrCON(3, I - 1) = "B" Then
                   If Val(sMatrCON(1, I - 1)) >= Val(sMatrCON(4, I - 1)) Then
                      sMin = sMatrCON(1, I - 1)
                   Else
                      sMin = sMatrCON(4, I - 1)
                   End If
                   If Val(sMatrCON(2, I - 1)) = 0 Then
                        sMax = sMatrCON(5, I - 1)
                   Else
                        If Val(sMatrCON(2, I - 1)) >= Val(sMatrCON(5, I - 1)) Then
                           sMax = sMatrCON(5, I - 1)
                        Else
                           sMax = sMatrCON(2, I - 1)
                        End If
                   End If
                   sFL = "A"
                Else
                   sFL = "A"
                   sMin = sMatrCON(1, I - 1)
                   sMax = sMatrCON(2, I - 1)
                End If
               
            Else
                  If sMatrCON(3, I - 1) = "A" Or sMatrCON(3, I - 1) = "B" Then
                     sFL = "A"
                     sMin = sMatrCON(4, I - 1)
                     sMax = sMatrCON(5, I - 1)
                  Else
                     sFL = ""
                     sMin = ""
                     sMax = ""
                  End If
                  
            End If
            With ss3
                .Row = I
                .Col = 2: .Text = sMin
                .Col = 3: .Text = sMax
                .Col = 5: .Text = sFL
            End With
            
         Next I
    End If
     
     Call subSpreadCheck1
     Call subSpreadERROR(ss3)
      With ss3
        For I = 1 To .MaxRows
            sRE = Gf_Get_Cell_Value(ss3, I, 4)
            sFL = Gf_Get_Cell_Value(ss3, I, 5)
            If sFL = "A" And sRE = "" Then
             .Col = 4
             .BackColor = RED
            End If
        Next I
      End With
    

End Sub

Private Sub subSpreadView2(ByVal strArr As Variant)

    Dim I As Integer
    Dim iRow As Integer
'    Dim sChem(34) As String
    Dim sChem(50) As String
    
    If UBound(strArr) < 104 Then Exit Sub
    
    sChem(0) = "C  "
    sChem(1) = "Mn "
    sChem(2) = "P  "
    sChem(3) = "S  "
    sChem(4) = "Si "
    sChem(5) = "Nb "
    sChem(6) = "Als"
    sChem(7) = "Alt"
    sChem(8) = "Ceq"
    sChem(9) = "Ni "
    sChem(10) = "Cr "
    sChem(11) = "Cu "
    sChem(12) = "Mo "
    sChem(13) = "V  "
    sChem(14) = "Ti "
    sChem(15) = "Pcm"
    sChem(16) = "W  "
    sChem(17) = "B  "
    sChem(18) = "Pb "
    sChem(19) = "Ca "
    sChem(20) = "N  "
    sChem(21) = "O  "
    sChem(22) = "H  "
    sChem(23) = "Zr "
    sChem(24) = "Mg "
    sChem(25) = "Sn "
    sChem(26) = "As "
    sChem(27) = "Co "
    sChem(28) = "Te "
    sChem(29) = "Bi "
    sChem(30) = "Sb "
    sChem(31) = "Zn "
    sChem(32) = "RE "
    sChem(33) = "Se "
    sChem(34) = "Ta "
    sChem(35) = "ComA"
    sChem(36) = "ComB"
    sChem(37) = "ComC"
    sChem(38) = "ComD"
    sChem(39) = "ComE"
    sChem(40) = "ComF"
    sChem(41) = "ComG"
    sChem(42) = "ComH"
    sChem(43) = "ComI"
    sChem(44) = "ComJ"
    sChem(45) = "ComK"
    sChem(46) = "ComL"
    sChem(47) = "ComN"
    sChem(48) = "ComP"
    sChem(49) = "ComV"
    sChem(50) = "Com1"

    For I = 0 To 50
        
        arrChem(0, I) = NullCheck(strArr(I, 0), "")
    
    Next I
    
    For I = 0 To 50
        
        arrChem(1, I) = NullCheck(strArr(I + 51, 0))
    
    Next I

    For I = 0 To 50
        
        arrChem(2, I) = NullCheck(strArr(I + 102, 0))
    
    Next I
    
    With SS2
    
        .MaxRows = 0
        .MaxRows = 51
    
        For I = 1 To 51
            .Row = I
            .Col = 1: .Text = sChem(I - 1)
            .Col = 2: .Text = arrChem(1, I - 1)
            .Col = 3: .Text = arrChem(0, I - 1)
            .Col = 4: .Text = arrChem(2, I - 1)
        Next I
          
    End With
    
    Call subSpreadCheck2
    Call subSpreadERROR(SS2)
End Sub

Private Sub subSpreadCheck2()
    
    Dim I As Long
    Dim j As Long
    
    j = 15
    With SS2
        
        For I = 16 To 51
                                    
            If (Gf_Get_Cell_Value(SS2, I, 4) = "" Or Gf_Get_Cell_Value(SS2, I, 4) = "0") _
               And (Gf_Get_Cell_Value(SS2, I, 2) = "0" And Gf_Get_Cell_Value(SS2, I, 3) = "0") Then
                .Row = I
                .RowHidden = True
            Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j
            End If
        Next I
                
    End With
    
End Sub

Private Sub subSpreadCheck1()
    
    Dim I As Long
    Dim j As Long
    
    With ss3
       
       For I = 1 To 143

           If Gf_Get_Cell_Value(ss3, I, 5) <> "A" And Gf_Get_Cell_Value(ss3, I, 5) <> "B" Then
               .Row = I
               .RowHidden = True
           Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j

           End If
            If Mid(Trim(TXT_STDSPEC), 1, 3) <> "API" And TXT_STDSPEC <> "GB/T9711.2-L450MB" Then
                If I = 20 Or I = 21 Or I = 22 Or I = 23 Or I = 24 _
                   Or I = 25 Or I = 26 Or I = 67 Or I = 68 Or I = 69 _
                   Or I = 70 Or I = 71 Or I = 72 Or I = 73 Then
                   .RowHidden = True
                End If
            End If
       Next I
                
    End With
End Sub
'Private Sub txt_SMP_CUT_LOC_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
''        If txt_SMP_NO = "" Then
''           MsgBox "请先输入取样号！", vbCritical, "系统提示信息"
''           txt_SMP_CUT_LOC = ""
''           Exit Sub
''        End If
'
'        DD.sWitch = "MS"
'        DD.sKey = "Q0042"
'        DD.rControl.Add Item:=txt_SMP_CUT_LOC
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'End Sub

'Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'
'    If Row < 1 Then Exit Sub
'
'    If Gf_Sc_Authority(sAuthority, "U") Then
'        With SS1
'            .Row = Row
'            .Col = 5
'            If .BackColor = &HFFFF& Then
'                .Col = 1
'                If .Text = "1" Then
'                    .Col = 5:   .Text = ""
'                    .Col = 0:   .Text = "Update"
'                Else
'                    .Col = 5:   .Text = "Y"
'                    .Col = 0:   .Text = ""
'                End If
'            Else
'                .Col = 1
'                If .Text = "1" Then
'                    .Col = 5:   .Text = "Y"
'                    .Col = 0:   .Text = "Update"
'                Else
'                    .Col = 5:   .Text = ""
'                    .Col = 0:   .Text = ""
'                End If
'            End If
'        End With
'    End If
'
'End Sub

'Private Sub txt_STDSPEC_Change()
'    If Trim(txt_STDSPEC.Text) = "" Then
'        txt_STDSPEC_NAME.Text = ""
'    End If
'
'End Sub
Private Sub subSpreadERROR(sPname As vaSpread)

    Dim I As Long
    Dim C_MAX, C_MIN, C_RESULT, C_FL As Variant

    With sPname

       If .MaxRows < 1 Then Exit Sub

       For I = 1 To .MaxRows
           .Row = I
           C_MIN = Val(Gf_Get_Cell_Value(sPname, I, 2))
           C_MAX = Val(Gf_Get_Cell_Value(sPname, I, 3))
           C_RESULT = Val(Gf_Get_Cell_Value(sPname, I, 4))
           If C_MIN <> 0 And C_MAX <> 0 Then
              If C_RESULT > C_MAX Or C_RESULT < C_MIN Then
                 Call Gp_Sp_CellColor(sPname, 4, I, RED)
              End If
           Else
              If C_MIN = 0 And C_MAX <> 0 Then
                 If C_RESULT > C_MAX Then
                    Call Gp_Sp_CellColor(sPname, 4, I, RED)
                 End If
              Else
                 If C_MIN <> 0 And C_MAX = 0 Then
                    If C_RESULT < C_MIN Then
                      Call Gp_Sp_CellColor(sPname, 4, I, RED)
                    End If
                 End If
              End If
           End If

       Next I

    End With

End Sub


Private Sub txt_PROD_CD_Change()
    Select Case txt_PROD_CD.Text
        Case "S", "s", "SL"
            txt_PROD_CD.Text = "SL"
        Case "P", "p", "PP"
            txt_PROD_CD.Text = "PP"
        Case "H", "h", "HC"
            txt_PROD_CD.Text = "HC"
        Case "", "**"
            txt_PROD_CD.Text = ""
        Case Else
            txt_PROD_CD.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
        End Select
End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_PROD_CD
    
        DD.nameType = "2"
    
        Call Gf_Common_DD(M_CN1, KeyCode)
    
    End If

End Sub

