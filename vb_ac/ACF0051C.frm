VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACF0051C 
   Caption         =   "生产简报_ACF0050C"
   ClientHeight    =   10395
   ClientLeft      =   285
   ClientTop       =   2325
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   15240
   Visible         =   0   'False
   WindowState     =   2  'Maximized
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
      Left            =   14040
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "ss1"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8460
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   14923
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACF0051C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   1245
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   15420
         _Version        =   393216
         _ExtentX        =   27199
         _ExtentY        =   2196
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
         MaxCols         =   5
         MaxRows         =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "ACF0051C.frx":00D2
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   1560
         Left            =   0
         TabIndex        =   2
         Top             =   1290
         Width           =   15420
         _Version        =   393216
         _ExtentX        =   27199
         _ExtentY        =   2752
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
         MaxCols         =   19
         MaxRows         =   4
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0051C.frx":05DF
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   1770
         Left            =   0
         TabIndex        =   8
         Top             =   6690
         Width           =   8115
         _Version        =   393216
         _ExtentX        =   14314
         _ExtentY        =   3122
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
         MaxCols         =   4
         MaxRows         =   4
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0051C.frx":1113
      End
      Begin FPSpread.vaSpread ss5 
         Height          =   1770
         Left            =   8160
         TabIndex        =   9
         Top             =   6690
         Width           =   7260
         _Version        =   393216
         _ExtentX        =   12806
         _ExtentY        =   3122
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
         MaxCols         =   7
         MaxRows         =   4
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0051C.frx":156A
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   1875
         Left            =   0
         TabIndex        =   10
         Top             =   2895
         Width           =   15420
         _Version        =   393216
         _ExtentX        =   27199
         _ExtentY        =   3307
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
         MaxCols         =   14
         MaxRows         =   7
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACF0051C.frx":1B1B
      End
      Begin FPSpread.vaSpread ss6 
         Height          =   1830
         Left            =   0
         TabIndex        =   11
         Top             =   4815
         Width           =   15420
         _Version        =   393216
         _ExtentX        =   27199
         _ExtentY        =   3228
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
         MaxRows         =   12
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         SpreadDesigner  =   "ACF0051C.frx":2448
      End
   End
   Begin InDate.ULabel ULabel3 
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   240
      Top             =   80
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "生产日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.UDate prod_date_from 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Tag             =   "INS_DATE"
      Top             =   80
      Width           =   1410
      _ExtentX        =   2487
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
   Begin InDate.UDate prod_date_to 
      Height          =   315
      Left            =   3030
      TabIndex        =   4
      Tag             =   "生产日期"
      Top             =   80
      Width           =   1410
      _ExtentX        =   2487
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
   Begin Threed.SSCommand Cmd_Edit 
      Height          =   360
      Left            =   8040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "更新数据"
   End
   Begin Threed.SSCommand BED_Cmd_Edit 
      Height          =   360
      Left            =   10400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "库存数据更新"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   2910
      TabIndex        =   5
      Top             =   200
      Width           =   90
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   15120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   15105
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "ACF0051C"
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
'-- Program ID        ACB1022C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2003.9.26
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim pColumn5 As New Collection      'Spread Primary Key Collection
Dim nColumn5 As New Collection      'Spread necessary Column Collection
Dim mColumn5 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn5 As New Collection      'Spread Insert Column Collection
Dim aColumn5 As New Collection      'Master -> Spread Column Collection
Dim lColumn5 As New Collection      'Spread Lock Column Collection

Dim pColumn6 As New Collection      'Spread Primary Key Collection
Dim nColumn6 As New Collection      'Spread necessary Column Collection
Dim mColumn6 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn6 As New Collection      'Spread Insert Column Collection
Dim aColumn6 As New Collection      'Master -> Spread Column Collection
Dim lColumn6 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim sc5 As New Collection           'Spread Collection
Dim sc6 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS2_PLT = 1
Const SS2_SLAB_WGT = 2
Const SS2_MILL_WGT = 3
Const SS2_BED_WGT = 4
Const SS2_UNBED_WGT = 5
Const SS2_UNBEDXAA_WGT = 6
Const SS2_UNBEDXAC_WGT = 7
Const SS2_UNBEDQ_WGT = 8
Const SS2_UNPLAN_WGT = 9
Const SS2_UNPLAN_RATE = 10
Const SS2_CONTRACT_WGT = 11
Const SS2_CONTRACT_RATE = 12
Const SS2_PLATE_RATE = 14
Const SS2_HW_WGT = 16
Const SS2_HW_RATE = 17
Const SS2_WIP_WGT = 18
Const SS2_HEAT_WGT = 19


Const SS3_PLT = 1
Const SS3_PLATE_HTM_WGT = 2
Const SS3_PLATE_WGT = 3
Const SS3_N_WGT = 4
Const SS3_T_WGT = 5
Const SS3_QT_WGT = 6
Const SS3_NT_WGT = 7
Const SS3_TTT_WGT = 8
Const SS3_NN_WGT = 9
Const SS3_TQT_WGT = 10
Const SS3_NNN_WGT = 11
Const SS3_QQT_WGT = 12
Const SS3_TT_WGT = 13
Const SS3_UNSAMED_WGT = 14



Dim sWgtLenFlag As String
Dim sQuery  As String

Private Sub Form_Define()


        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(prod_date_from, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(prod_date_to, "p", "n", "", " ", "r", " ", "", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "", " ", " ", "", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACF0050C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    
'    sc1.Item("Spread").Col = 0
'    sc1.Item("Spread").Row = 0
'    sc1.Item("Spread").Text = "◎"


    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACF0050C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Call Gp_Sp_ColHidden(ss2, 19, True)
    
    
    
'    sc2.Item("Spread").Col = 0
'    sc2.Item("Spread").Row = 0
'    sc2.Item("Spread").Text = "◎"

     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)

    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="ACF0050C.P_REFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    
     Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)


    'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="ACF0050C.P_REFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc4, Key:="Sc4"
    
     Call Gp_Sp_Collection(ss5, 1, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 2, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 3, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 4, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 5, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 6, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 7, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     
    'Spread_Collection
    sc5.Add Item:=ss5, Key:="Spread"
    sc5.Add Item:="ACF0050C.P_REFER5", Key:="P-R"
    sc5.Add Item:=pColumn5, Key:="pColumn"
    sc5.Add Item:=nColumn5, Key:="nColumn"
    sc5.Add Item:=aColumn5, Key:="aColumn"
    sc5.Add Item:=mColumn5, Key:="mColumn"
    sc5.Add Item:=iColumn5, Key:="iColumn"
    sc5.Add Item:=lColumn5, Key:="lColumn"
    sc5.Add Item:=1, Key:="First"
    sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc5, Key:="Sc5"
    
    Call Gp_Sp_Collection(ss6, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss6, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss6, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss6, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss6, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss6, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss6, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss6, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc6.Add Item:=ss6, Key:="Spread"
    sc6.Add Item:="ACF0050C.P_REFER6", Key:="P-R"
    sc6.Add Item:=pColumn6, Key:="pColumn"
    sc6.Add Item:=nColumn6, Key:="nColumn"
    sc6.Add Item:=aColumn6, Key:="aColumn"
    sc6.Add Item:=mColumn6, Key:="mColumn"
    sc6.Add Item:=iColumn6, Key:="iColumn"
    sc6.Add Item:=lColumn6, Key:="lColumn"
    sc6.Add Item:=1, Key:="First"
    sc6.Add Item:=ss6.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc6, Key:="Sc6"
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    
'    Sc3.Item("Spread").Col = 0
'    Sc3.Item("Spread").Row = 0
'    Sc3.Item("Spread").Text = "◎"
    
        
End Sub

Private Sub BED_Cmd_Edit_Click()

  Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
          
    If Trim(prod_date_to) = "" Then
        Call Gp_MsgBoxDisplay(prod_date_to.Tag + "必须输入")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACF0051P ('" + Trim(Format(prod_date_from.Text, "YYYYMMDD")) + "','" + Trim(Format(prod_date_to.Text, "YYYYMMDD")) + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
            
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("更新成功..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")


End Sub

Private Sub Cmd_Edit_Click()
'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
          
    If Trim(prod_date_to) = "" Then
        Call Gp_MsgBoxDisplay(prod_date_to.Tag + "必须输入")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACF0050P ('" + Trim(Format(prod_date_from.Text, "YYYYMMDD")) + "','" + Trim(Format(prod_date_to.Text, "YYYYMMDD")) + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
            
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("更新成功..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste

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
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(sc4.Item("Spread"), False)
    Call Gp_Sp_Setting(sc5.Item("Spread"), False)
    Call Gp_Sp_Setting(sc6.Item("Spread"), False)
    
    'Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
'    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
     'Call Gp_Sp_ReadOnlySet(Sc4.Item("Spread"))
     'Call Gp_Sp_ReadOnlySet(Sc5.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(sc4)
    Call Gf_Sp_Cls(sc5)
    Call Gf_Sp_Cls(sc6)
    
    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc5.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc6.Item("Spread"), "C-System.INI", Me.Name)
    
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 5)
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 7)
    'Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 14)
    'Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 15)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc5.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc6.Item("Spread"), "C-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
      
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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
    
    Set iColumn5 = Nothing
    Set pColumn5 = Nothing
    Set lColumn5 = Nothing
    Set nColumn5 = Nothing
    Set mColumn5 = Nothing
    Set aColumn5 = Nothing
    
    Set iColumn6 = Nothing
    Set pColumn6 = Nothing
    Set lColumn6 = Nothing
    Set nColumn6 = Nothing
    Set mColumn6 = Nothing
    Set aColumn6 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set sc4 = Nothing
    Set sc5 = Nothing
    Set sc6 = Nothing
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Or Gf_Sp_Cls(Sc3) Or Gf_Sp_Cls(sc4) Or Gf_Sp_Cls(sc5) Or Gf_Sp_Cls(sc6) Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rContro1(1).SetFocus
   
        End If
    End If
    
End Sub

Public Sub Form_Ref()


Dim SLAB_WGT        As Double
Dim MILL_WGT        As Double
Dim BED_WGT         As Double
Dim UNPLAN_WGT      As Double
Dim UNPLAN_RATE As Double
Dim PLATE_RATE      As Double
Dim HW_WGT          As Double
Dim HW_RATE         As Double
Dim WIP_WGT         As Double
Dim HEAT_WGT        As Double
Dim CONTRACT_WGT    As Double
Dim CONTRACT_RATE   As Double

Dim iCount          As Integer

Dim PLATE_HTM_WGT  As Double
Dim UNSAMED_WGT    As Double
Dim PLATE_WGT         As Double
Dim N_WGT             As Double
Dim T_WGT             As Double
Dim QT_WGT            As Double
Dim NT_WGT            As Double
Dim TTT_WGT           As Double
Dim NN_WGT            As Double
Dim TQT_WGT           As Double
Dim NNN_WGT           As Double
Dim QQT_WGT           As Double
Dim TT_WGT            As Double
Dim UNBED_WGT         As Double
Dim UNBEDXAA_WGT      As Double
Dim UNBEDXAC_WGT      As Double
Dim UNBEDQ_WGT        As Double


         Call Gf_Sp_Cls(sc2)
         Call Gf_Sp_Cls(Sc3)
         Call Gf_Sp_Cls(sc1)
         Call Gf_Sp_Cls(sc4)
         Call Gf_Sp_Cls(sc5)
         Call Gf_Sp_Cls(sc6)
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rContro1(1).SetFocus
   
       
    

    
    'If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
        ss1.OperationMode = OperationModeNormal
        'Call Gp_Sp_BlockColor(ss1, 7, 7, 1, ss1.MaxRows)
        Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss2.OperationMode = OperationModeNormal
        
        Call Gf_Sp_Cls(Sc3)
        Call Gf_Sp_Refer(M_CN1, Sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss3.OperationMode = OperationModeNormal
        
        Call Gf_Sp_Refer(M_CN1, sc4, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss4.OperationMode = OperationModeNormal
        
         Call Gf_Sp_Refer(M_CN1, sc5, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss5.OperationMode = OperationModeNormal
        
           Call Gf_Sp_Refer(M_CN1, sc6, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss5.OperationMode = OperationModeNormal
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        MDIMain.MenuTool.Buttons(4).Enabled = True
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        
    
        
        With ss2
        If .MaxRows < 1 Then
           Exit Sub
        End If
        .MaxRows = .MaxRows + 1
        For iCount = 1 To .MaxRows - 1
        .ROW = iCount
        '坯料重量
         .Col = SS2_SLAB_WGT:                SLAB_WGT = SLAB_WGT + Val(.Text)
        '轧制产量
         .Col = SS2_MILL_WGT:                MILL_WGT = MILL_WGT + Val(.Text)
        '入库产量
        .Col = SS2_BED_WGT:                  BED_WGT = BED_WGT + Val(.Text)
        '非计划量
        .Col = SS2_UNPLAN_WGT:               UNPLAN_WGT = UNPLAN_WGT + Val(.Text)
        '合同附带量
        .Col = SS2_CONTRACT_WGT:             CONTRACT_WGT = CONTRACT_WGT + Val(.Text)
        '非计划率
        '.Col = SS2_UNPLAN_RATE:       UNPLAN_RATE = Val(.Text):                  UNPLAN_RATE = UNPLAN_RATE + Val(.Text)
        
        '成材率
        '.Col = SS2_PLATE_RATE:        PLATE_RATE = Val(.Text):                    PLATE_RATE = PLATE_RATE + Val(.Text)
        '热装热送量
        .Col = SS2_HW_WGT:                   HW_WGT = HW_WGT + Val(.Text)
        '热装热送率
        '.Col = SS2_HW_RATE:           HW_RATE = Val(.Text):                          HW_RATE = HW_RATE + Val(.Text)
        '在制品量
        .Col = SS2_WIP_WGT:                  WIP_WGT = WIP_WGT + Val(.Text)
        .Col = SS2_HEAT_WGT:                 HEAT_WGT = HEAT_WGT + Val(.Text)
        '未入库量
        .Col = SS2_UNBED_WGT:          UNBED_WGT = UNBED_WGT + Val(.Text)
        .Col = SS2_UNBEDXAA_WGT:       UNBEDXAA_WGT = UNBEDXAA_WGT + Val(.Text)
        .Col = SS2_UNBEDXAC_WGT:       UNBEDXAC_WGT = UNBEDXAC_WGT + Val(.Text)
        .Col = SS2_UNBEDQ_WGT:         UNBEDQ_WGT = UNBEDQ_WGT + Val(.Text)
        
        Next iCount
        
      .ROW = .MaxRows
        
.Col = SS2_PLT:                      .Text = "合计"
.Col = SS2_SLAB_WGT:                 .Text = SLAB_WGT
.Col = SS2_MILL_WGT:                 .Text = MILL_WGT
.Col = SS2_BED_WGT:                  .Text = BED_WGT
.Col = SS2_UNPLAN_WGT:               .Text = UNPLAN_WGT
.Col = SS2_CONTRACT_WGT:             .Text = CONTRACT_WGT
.Col = SS2_UNBED_WGT:                .Text = UNBED_WGT
.Col = SS2_UNBEDXAA_WGT:             .Text = UNBEDXAA_WGT
.Col = SS2_UNBEDXAC_WGT:             .Text = UNBEDXAC_WGT
.Col = SS2_UNBEDQ_WGT:               .Text = UNBEDQ_WGT

 If BED_WGT = 0 Then
 .Col = SS2_UNPLAN_RATE:             .Text = 0
 .Col = SS2_CONTRACT_RATE:           .Text = 0
 Else
.Col = SS2_UNPLAN_RATE:              .Text = UNPLAN_WGT / BED_WGT * 100
.Col = SS2_CONTRACT_RATE:            .Text = CONTRACT_WGT / BED_WGT * 100
End If

 If SLAB_WGT = 0 Then
 .Col = SS2_PLATE_RATE:             .Text = 0
 Else
.Col = SS2_PLATE_RATE:               .Text = MILL_WGT / SLAB_WGT * 100
End If

.Col = SS2_HW_WGT:                   .Text = HW_WGT

If SLAB_WGT = 0 Then
 .Col = SS2_HW_RATE:                 .Text = 0
 Else
.Col = SS2_HW_RATE:                  .Text = HW_WGT / SLAB_WGT * 100
End If

.Col = SS2_WIP_WGT:                  .Text = WIP_WGT

End With

With ss3
        If .MaxRows < 1 Then
           Exit Sub
        End If
        .MaxRows = .MaxRows + 1
        .ROW = 1
        .Col = SS3_PLATE_WGT:             PLATE_WGT = Val(.Text)
        .Col = SS3_N_WGT:                 N_WGT = Val(.Text)
        .Col = SS3_T_WGT:                 T_WGT = Val(.Text)
        .Col = SS3_QT_WGT:                QT_WGT = Val(.Text)
        .Col = SS3_NT_WGT:                NT_WGT = Val(.Text)
        .Col = SS3_TTT_WGT:               TTT_WGT = Val(.Text)
        .Col = SS3_NN_WGT:                NN_WGT = Val(.Text)
        .Col = SS3_TQT_WGT:               TQT_WGT = Val(.Text)
        .Col = SS3_NNN_WGT:               NNN_WGT = Val(.Text)
        .Col = SS3_QQT_WGT:               QQT_WGT = Val(.Text)
        .Col = SS3_TT_WGT:               TT_WGT = Val(.Text)
        .Col = SS3_UNSAMED_WGT:           UNSAMED_WGT = Val(.Text)

        
        For iCount = 1 To .MaxRows - 1
        
        .ROW = iCount
        '过钢量
        .Col = SS3_PLATE_HTM_WGT:      PLATE_HTM_WGT = PLATE_HTM_WGT + Val(.Text)
        '其中与录单热处理方式不符量
        '.Col = SS3_UNSAMED_WGT:       UNSAMED_WGT = Val(.Text):                  UNSAMED_WGT = UNSAMED_WGT + Val(.Text)
        Next iCount
        
        .ROW = .MaxRows
        
                                                              
                                                              
.Col = SS3_PLT:                      .Text = "合计"
.Col = SS3_PLATE_HTM_WGT:            .Text = PLATE_HTM_WGT
.Col = SS3_UNSAMED_WGT:              .Text = UNSAMED_WGT

.Col = SS3_PLATE_WGT:             .Text = PLATE_WGT
.Col = SS3_N_WGT:                 .Text = N_WGT
.Col = SS3_T_WGT:                 .Text = T_WGT
.Col = SS3_QT_WGT:                .Text = QT_WGT
.Col = SS3_NT_WGT:                .Text = NT_WGT
.Col = SS3_TTT_WGT:               .Text = TTT_WGT
.Col = SS3_NN_WGT:                .Text = NN_WGT
.Col = SS3_TQT_WGT:               .Text = TQT_WGT
.Col = SS3_NNN_WGT:               .Text = NNN_WGT
.Col = SS3_QQT_WGT:               .Text = QQT_WGT
.Col = SS3_TT_WGT:                .Text = TT_WGT

.ROW = 1

.Col = SS3_UNSAMED_WGT:           .Text = ""
                                            
.Col = SS3_PLATE_WGT:             .Text = ""
.Col = SS3_N_WGT:                 .Text = ""
.Col = SS3_T_WGT:                 .Text = ""
.Col = SS3_QT_WGT:                .Text = ""
.Col = SS3_NT_WGT:                .Text = ""
.Col = SS3_TTT_WGT:               .Text = ""
.Col = SS3_NN_WGT:                .Text = ""
.Col = SS3_TQT_WGT:               .Text = ""
.Col = SS3_NNN_WGT:               .Text = ""
.Col = SS3_QQT_WGT:               .Text = ""
.Col = SS3_TT_WGT:                .Text = ""

End With
        'txt_charge_no.Text = ""
End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Form_Exc()


 If txt_shape.Text = "ss1" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf txt_shape.Text = "ss2" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf txt_shape.Text = "ss3" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf txt_shape.Text = "ss4" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc4")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf txt_shape.Text = "ss5" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc5")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
     ElseIf txt_shape.Text = "ss6" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc6")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub




Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
txt_shape.Text = "ss1"
End Sub



Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
txt_shape.Text = "ss2"
End Sub


Private Sub ss3_Click(ByVal Col As Long, ByVal ROW As Long)
txt_shape.Text = "ss3"
End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal ROW As Long)
txt_shape.Text = "ss4"
End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal ROW As Long)
txt_shape.Text = "ss5"
End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal ROW As Long)
txt_shape.Text = "ss6"
End Sub
