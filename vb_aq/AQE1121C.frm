VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQE1121C 
   Caption         =   "理化室取样性能查询_AQE1121C"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_PLT 
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
      Height          =   300
      Left            =   4590
      MaxLength       =   2
      TabIndex        =   11
      Top             =   150
      Width           =   420
   End
   Begin VB.TextBox txt_PLT_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5085
      TabIndex        =   10
      Top             =   150
      Width           =   2430
   End
   Begin VB.TextBox txt_STDSPEC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   13575
      MaxLength       =   18
      TabIndex        =   4
      Top             =   150
      Width           =   2295
   End
   Begin VB.TextBox txt_STDSPEC_NAME 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   15915
      MaxLength       =   18
      TabIndex        =   3
      Top             =   150
      Width           =   570
   End
   Begin InDate.UDate dtp_date_t 
      Height          =   315
      Left            =   10785
      TabIndex        =   2
      Top             =   150
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.UDate dtp_date_f 
      Height          =   315
      Left            =   9000
      TabIndex        =   1
      Top             =   150
      Width           =   1470
      _ExtentX        =   2593
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
   Begin VB.TextBox txt_HEAT_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1215
      MaxLength       =   14
      TabIndex        =   0
      Top             =   135
      Width           =   1965
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "冶炼炉号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   2
      Left            =   7905
      Top             =   150
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Caption         =   "生产日期"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   12600
      Top             =   150
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "标准编号"
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
      ForeColor       =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7935
      Left            =   0
      TabIndex        =   5
      Top             =   645
      Width           =   18285
      _ExtentX        =   32253
      _ExtentY        =   13996
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AQE1121C.frx":0000
      Begin FPSpread.vaSpread SS2 
         Height          =   7935
         Left            =   10380
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   3450
         _Version        =   393216
         _ExtentX        =   6085
         _ExtentY        =   13996
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
         SpreadDesigner  =   "AQE1121C.frx":0092
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   7935
         Left            =   13920
         TabIndex        =   7
         Top             =   0
         Width           =   4365
         _Version        =   393216
         _ExtentX        =   7699
         _ExtentY        =   13996
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
         SpreadDesigner  =   "AQE1121C.frx":0412
      End
      Begin FPSpread.vaSpread SS1 
         Height          =   7935
         Left            =   3810
         TabIndex        =   8
         Top             =   0
         Width           =   6480
         _Version        =   393216
         _ExtentX        =   11430
         _ExtentY        =   13996
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
         MaxCols         =   10
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQE1121C.frx":0787
      End
      Begin FPSpread.vaSpread SS4 
         Height          =   7935
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3720
         _Version        =   393216
         _ExtentX        =   6562
         _ExtentY        =   13996
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
         MaxCols         =   3
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQE1121C.frx":0E98
      End
   End
   Begin InDate.ULabel lab_COLOR_STROKE 
      Height          =   315
      Left            =   120
      Top             =   8760
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   556
      Caption         =   "订单备注："
      Alignment       =   0
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
      ForeColor       =   255
   End
   Begin InDate.ULabel lab_MATR_FL 
      Height          =   315
      Index           =   0
      Left            =   13560
      Top             =   8760
      Width           =   4700
      _ExtentX        =   8281
      _ExtentY        =   556
      Caption         =   " 产品力学性能："
      Alignment       =   0
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
      ForeColor       =   255
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   6
      Left            =   3600
      Top             =   150
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "生产厂"
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
      ForeColor       =   0
   End
   Begin FPSpread.vaSpread SS5 
      Height          =   1335
      Left            =   18360
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   1920
      _Version        =   393216
      _ExtentX        =   3387
      _ExtentY        =   2355
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
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE1121C.frx":12B9
   End
End
Attribute VB_Name = "AQE1121C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   报表管理
'-- Program Name      理化室取样性能查询
'-- Program ID        AQE1121C
'-- Coder             LIU XIANG
'-- Date              2011.11.25
'-- Description       理化室取样性能查询
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
Public sPLT_Authority As String     'Active User Plant Authority Setting

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection
Dim sc3 As New Collection
Dim sc4 As New Collection
Dim sc5 As New Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim arrChem(3, 35) As String

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    Call Gp_Ms_Collection(txt_HEAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(dtp_date_f, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(dtp_date_t, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=cControl1, Key:="cControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"
    
    Call Gp_Ms_Collection(txt_HEAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
       
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     
     'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="AQE1121C.P_REFER", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc4, Key:="Sc"
  
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
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQE1121C.P_SREFER_0", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Proc_Sc.Add Item:=Sc1, Key:="Sc1"
'     Call SS1.AddCellSpan(5, 0, 1, 2)

      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     
     'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQE1121C.P_SREFER_1", Key:="P-R"
    sc2.Add Item:=pColumn12, Key:="pColumn"
    sc2.Add Item:=nColumn12, Key:="nColumn"
    sc2.Add Item:=aColumn12, Key:="aColumn"
    sc2.Add Item:=mColumn12, Key:="mColumn"
    sc2.Add Item:=iColumn12, Key:="iColumn"
    sc2.Add Item:=lColumn12, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     
     'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQE1121C.P_SREFER_2", Key:="P-R"
    sc3.Add Item:=pColumn13, Key:="pColumn"
    sc3.Add Item:=nColumn13, Key:="nColumn"
    sc3.Add Item:=aColumn13, Key:="aColumn"
    sc3.Add Item:=mColumn13, Key:="mColumn"
    sc3.Add Item:=iColumn13, Key:="iColumn"
    sc3.Add Item:=lColumn13, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
        
    Call Gp_Sp_Collection(ss5, 1, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 2, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 3, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    
     'Spread_Collection
    sc5.Add Item:=ss5, Key:="Spread"
    sc5.Add Item:="AQE1121C.P_SREFER_5", Key:="P-R"
    sc5.Add Item:=pColumn5, Key:="pColumn"
    sc5.Add Item:=nColumn5, Key:="nColumn"
    sc5.Add Item:=aColumn5, Key:="aColumn"
    sc5.Add Item:=mColumn5, Key:="mColumn"
    sc5.Add Item:=iColumn5, Key:="iColumn"
    sc5.Add Item:=lColumn5, Key:="lColumn"
    sc5.Add Item:=1, Key:="First"
    sc5.Add Item:=ss5.MaxCols, Key:="Last"

    
    
    'Call Gp_Sp_ColHidden(ss1, 6, True)
    'Call Gp_Sp_ColHidden(ss1, 7, True)
    Call Gp_Sp_ColHidden(ss2, 0, True)
    Call Gp_Sp_ColHidden(ss3, 0, True)
    Call Gp_Sp_ColHidden(ss3, 5, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, 1, ss1.MaxRows, , &HFFFF&)


End Sub




Private Sub MenuToolSet()
     
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    MDIMain.MenuTool.Buttons(14).Enabled = True    'EXCEL
'    MDIMain.MenuTool.Buttons(14).Enabled = False
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet
    txt_plt.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        Case "txt_PLT"                     '工厂
            sCode = "C0001"
            Set oCodeName = txt_PLT_NAME
        Case "txt_SMP_CUT_LOC"             '取样位置
            sCode = "Q0042"
            'Set oCodeName = txt_SMP_CUT_LOC_NAME
        Case "txt_STDSPEC"                 '标准
            sCode = "STDSPEC"
            Set oCodeName = txt_STDSPEC_NAME
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing

Err_Track:
    
    Set oCodeName = Nothing

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    sPLT_Authority = Gf_PLT_Authority(Me.Name)
    If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
       txt_plt.Text = sPLT_Authority
    Else
       txt_plt.Text = ""
    End If
    
    
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuToolSet

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(ss1)
    Call Gp_Sp_Setting(ss2)
    Call Gp_Sp_Setting(ss3)
    Call Gp_Sp_Setting(ss4)
    Call Gp_Sp_ReadOnlySet(ss1)
    Call Gp_Sp_ReadOnlySet(ss2)
    Call Gp_Sp_ReadOnlySet(ss3)
    Call Gp_Sp_ReadOnlySet(ss4)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss4, "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss2, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss3, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss4, "Q-System.INI", Me.Name)
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set sc4 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(Sc1)
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(sc3)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        
        If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
           txt_plt.Text = sPLT_Authority
        Else
           txt_plt.Text = ""
        End If
        
    End If
    
    txt_PLT_NAME = ""
    txt_STDSPEC_NAME.Text = ""
    txt_HEAT_NO.Text = ""

End Sub

Public Sub Form_Ref()
    On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_Refer(M_CN1, sc4, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss4.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuToolSet
    End If
    Call Gf_Sp_Refer(M_CN1, sc5, Mc1, Mc1("nControl"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    

Refer_Err:
End Sub

Public Sub Form_Pro()

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
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

Public Sub Form_Exc()

    'Call Gf_Sp_Refer(M_CN1, sc5, Mc1, Mc1("nControl"))
    Call Gp_Sp_Excel(Me, sc5("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    'Call Gp_Sp_Excel(Me, Proc_Sc("SC")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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
    Dim AdoRs           As adodb.Recordset
    Dim ArrayRecords    As Variant
    Dim arr             As Variant
    Dim SMP_NO, smp_loc As Variant
    Dim s_ORD_NO, s_ORD_ITEM As String
    Dim s_MATR_FL  As String
    
    Dim s_COLOR_STROKE  As String
    
    s_MATR_FL = "Y"
    
    s_COLOR_STROKE = ""
    
  
    
    
 On Error GoTo Error_Rtn
    
    Call Gp_Sp_Sort(Sc1("Spread"), Col, Row)

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
        
    With ss1
        .Col = 1
        .Row = .ActiveRow
        SMP_NO = .Text
        .Col = 2
        smp_loc = .Text
        .Col = 3
        txt_STDSPEC = .Text
        .Col = 7
        s_ORD_NO = .Text
        .Col = 8
        s_ORD_ITEM = .Text
    End With

    
    ss2.MaxRows = 0
    ss3.MaxRows = 0
    
    ss1.ReDraw = False
    ss2.ReDraw = False
    ss3.ReDraw = False
    
    Set AdoRs = New adodb.Recordset
    sQuery = "SELECT MATR_FL,COLOR_STROKE,INSP_CD  FROM BP_ORDER_ITEM WHERE ORD_NO = " + "'" + s_ORD_NO + "'"
    sQuery = sQuery + " AND ORD_ITEM = " + "'" + s_ORD_ITEM + "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not (AdoRs.BOF And AdoRs.EOF) And IsNull(AdoRs.Fields(0).Value) = False Then
        'ArrayRecords = AdoRs.GetRows
        s_MATR_FL = AdoRs.Fields(0).Value
    Else
        s_MATR_FL = "Y"
    End If
      
    If IsNull(AdoRs.Fields(0).Value) = False Then
'        s_COLOR_STROKE = AdoRs.Fields(1).Value
        lab_COLOR_STROKE.Caption = " 订单备注： " + s_ORD_NO + "-" + s_ORD_ITEM + ":" + AdoRs.Fields(1).Value
        If IsNull(AdoRs.Fields(2).Value) = False Then
            lab_COLOR_STROKE.Caption = lab_COLOR_STROKE.Caption + "  见证机关：" + AdoRs.Fields(2).Value
        End If
    End If
    

    sQuery = "{call AQC0040C.P_SREFER_1('" + Trim(SMP_NO) + "','" + Trim(smp_loc) + "')}"
    AdoRs.Close

    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView2(ArrayRecords)
        Erase ArrayRecords
    End If


'    Call Gp_Sp_EvenRowBackcolor(ss2)
        
    If s_MATR_FL = "N" Or s_MATR_FL = "n" Then
        lab_MATR_FL.Item(0).Caption = " 产品力学性能： " + " 不保证 "
    Else
        lab_MATR_FL.Item(0).Caption = " 产品力学性能： " + " 保证 "
    End If
    
   
    sQuery = "{call AQC0040C.P_SREFER_2('" + Trim(SMP_NO) + "','" + Trim(smp_loc) + "')}"
                    
    AdoRs.Close
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView1(ArrayRecords)
        Erase ArrayRecords
    Else                                '如果还没性能结果ss3中没有行，后面的subSpreadView3会造成ss3标题列显示错误，所以在此退出
        Exit Sub
    End If
     
    sQuery = "{call AQC0040C.P_SREFER_3('" + Trim(SMP_NO) + "')}"
    
    AdoRs.Close
                    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView3(ArrayRecords)
        Erase ArrayRecords
    End If
    

'    Call Gp_Sp_EvenRowBackcolor(ss3)
    
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True
       
    Exit Sub
    
Error_Rtn:
    
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    Screen.MousePointer = vbDefault
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True

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
Private Sub subSpreadView1(ByVal strArr As Variant)

    Dim i           As Integer
    Dim iRow        As Integer
    Dim sMatr(166)   As String
    
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
   
    sMatr(19) = "冲击断面纤维率实绩平均                 "
    sMatr(20) = "冲击断面纤维率实绩 1                   "
    sMatr(21) = "冲击断面纤维率实绩 2                   "
    sMatr(22) = "冲击断面纤维率实绩 3                   "
    sMatr(23) = "冲击断面纤维率实绩 4                   "
    sMatr(24) = "冲击断面纤维率实绩 5                   "
    sMatr(25) = "冲击断面纤维率实绩 6                   "
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
    sMatr(70) = "追加冲击断面纤维率实绩 1               "
    sMatr(71) = "追加冲击断面纤维率实绩 2               "
    sMatr(72) = "追加冲击断面纤维率实绩 3               "
    sMatr(73) = "追加冲击断面纤维率实绩 4               "
    sMatr(74) = "追加冲击断面纤维率实绩 5               "
    sMatr(75) = "追加冲击断面纤维率实绩 6               "
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

'louyanan 20101121 start

   sMatr(136) = "拉伸厚度方向断面收缩率实绩1"
   sMatr(137) = "拉伸厚度方向断面收缩率实绩2"
   sMatr(138) = "拉伸厚度方向断面收缩率实绩3"
   sMatr(139) = "拉伸厚度方向断面收缩率实绩平均"
   sMatr(140) = "高温拉伸厚度方向断面收缩率实绩1"
   sMatr(141) = "高温拉伸厚度方向断面收缩率实绩2"
   sMatr(142) = "高温拉伸厚度方向断面收缩率实绩3"
   sMatr(143) = "高温拉伸厚度方向断面收缩率实绩平均"
   sMatr(144) = "冲击侧向膨胀值实绩1"
   sMatr(145) = "冲击侧向膨胀值实绩2"
   sMatr(146) = "冲击侧向膨胀值实绩3"
   sMatr(147) = "冲击侧向膨胀值实绩4"
   sMatr(148) = "冲击侧向膨胀值实绩5"
   sMatr(149) = "冲击侧向膨胀值实绩6"
   sMatr(150) = "冲击侧向膨胀值实绩平均"
   sMatr(151) = "追加冲击侧向膨胀值实绩1"
   sMatr(152) = "追加冲击侧向膨胀值实绩2"
   sMatr(153) = "追加冲击侧向膨胀值实绩3"
   sMatr(154) = "追加冲击侧向膨胀值实绩4"
   sMatr(155) = "追加冲击侧向膨胀值实绩5"
   sMatr(156) = "追加冲击侧向膨胀值实绩6"
   sMatr(157) = "追加冲击侧向膨胀值实绩平均"
   sMatr(158) = "NDT重力撕裂实绩"
'edit by gengxueyu 20110212 for kangda start
   sMatr(159) = "均匀变形伸长率UEL"
   sMatr(160) = "追加均匀变形伸长率UEL"
   sMatr(161) = "追加应力比项目1"
   sMatr(162) = "追加应力比项目2"
   sMatr(163) = "追加应力比项目3"
   sMatr(164) = "追加应力比项目4"
   sMatr(165) = "追加应力比项目5"
'edit by gengxueyu 20110212 for kangda end
   
   
  
  
  
  
    With ss3
        .MaxRows = 166
    
        For i = 1 To 166
            .Row = i
            .Col = 1: .Text = sMatr(i - 1)
        Next i
                
        For i = 1 To UBound(strArr, 1) + 1
        
            .Row = i: .Col = 4
            .Text = NullCheck(strArr(i - 1, 0), "")
            
        Next i
    End With

End Sub

Private Sub subSpreadView3(ByVal strArr As Variant)

    Dim i                     As Integer
    Dim iRow                  As Integer
    Dim sMatr(3, 166)         As Variant
    Dim sMatrCON(6, 166)      As Variant
    Dim sMin, sMax, sFL, sRE  As Variant
    
    If UBound(strArr, 2) < 0 Then Exit Sub
      
    If UBound(strArr, 2) = 0 Then
        For i = 0 To 165
            sMatr(0, i) = NullCheck(strArr(i, 0), "")
        Next i
        
        For i = 0 To 165
            sMatr(1, i) = NullCheck(strArr(i + 166, 0))
        Next i
    
        For i = 0 To 165
            sMatr(2, i) = NullCheck(strArr(i + 332, 0))
        Next i
        
        
        With ss3
                
            For i = 1 To 166
                .Row = i
                .Col = 2: .Text = sMatr(1, i - 1)
                .Col = 3: .Text = sMatr(2, i - 1)
                .Col = 5: .Text = sMatr(0, i - 1)
            Next i
         End With
    End If
     
    If UBound(strArr, 2) = 1 Then
        For i = 0 To 165
            sMatrCON(0, i) = NullCheck(strArr(i, 0), "")
            sMatrCON(3, i) = NullCheck(strArr(i, 1), "")
        Next i
        
        For i = 0 To 165
            sMatrCON(1, i) = NullCheck(strArr(i + 166, 0))
            sMatrCON(4, i) = NullCheck(strArr(i + 166, 1))
        Next i
    
        For i = 0 To 165
            sMatrCON(2, i) = NullCheck(strArr(i + 332, 0))
            sMatrCON(5, i) = NullCheck(strArr(i + 332, 1))
        Next i
        
            
        For i = 1 To 165
            If sMatrCON(0, i - 1) = "A" Or sMatrCON(0, i - 1) = "B" Then
                If sMatrCON(3, i - 1) = "A" Or sMatrCON(3, i - 1) = "B" Then
                   If Val(sMatrCON(1, i - 1)) >= Val(sMatrCON(4, i - 1)) Then
                      sMin = sMatrCON(1, i - 1)
                   Else
                      sMin = sMatrCON(4, i - 1)
                   End If
                   If Val(sMatrCON(2, i - 1)) = 0 Then
                        sMax = sMatrCON(5, i - 1)
                   Else
                        If Val(sMatrCON(2, i - 1)) >= Val(sMatrCON(5, i - 1)) Then
                           sMax = sMatrCON(5, i - 1)
                        Else
                           sMax = sMatrCON(2, i - 1)
                        End If
                   End If
                   sFL = "A"
                Else
                   sFL = "A"
                   sMin = sMatrCON(1, i - 1)
                   sMax = sMatrCON(2, i - 1)
                End If
               
            Else
                  If sMatrCON(3, i - 1) = "A" Or sMatrCON(3, i - 1) = "B" Then
                     sFL = "A"
                     sMin = sMatrCON(4, i - 1)
                     sMax = sMatrCON(5, i - 1)
                  Else
                     sFL = ""
                     sMin = ""
                     sMax = ""
                  End If
                  
            End If
            With ss3
                .Row = i
                .Col = 2: .Text = sMin
                .Col = 3: .Text = sMax
                .Col = 5: .Text = sFL
            End With
            
         Next i
    End If
     
     Call subSpreadCheck1
     Call subSpreadERROR(ss3)
      With ss3
        For i = 1 To .MaxRows
            sRE = Gf_Get_Cell_Value(ss3, i, 4)
            sFL = Gf_Get_Cell_Value(ss3, i, 5)
            If sFL = "A" And sRE = "" Then
             .Col = 4
             .BackColor = RED
            End If
        Next i
      End With
    

End Sub

Private Sub subSpreadView2(ByVal strArr As Variant)

    Dim i As Integer
    Dim iRow As Integer
    Dim sChem(34) As String
    
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

    For i = 0 To 34
        
        arrChem(0, i) = NullCheck(strArr(i, 0), "")
    
    Next i
    
    For i = 0 To 34
        
        arrChem(1, i) = NullCheck(strArr(i + 35, 0))
    
    Next i

    For i = 0 To 34
        
        arrChem(2, i) = NullCheck(strArr(i + 70, 0))
    
    Next i
    
    With ss2
    
        .MaxRows = 0
        .MaxRows = 35
    
        For i = 1 To 35
            .Row = i
            .Col = 1: .Text = sChem(i - 1)
            .Col = 2: .Text = arrChem(1, i - 1)
            .Col = 3: .Text = arrChem(0, i - 1)
            .Col = 4: .Text = arrChem(2, i - 1)
        Next i
          
    End With
    
    Call subSpreadCheck2
    Call subSpreadERROR(ss2)
End Sub

Private Sub subSpreadCheck2()
    
    Dim i As Long
    Dim j As Long
    
    j = 15
    With ss2
        
        For i = 16 To 35
                                    
            If (Gf_Get_Cell_Value(ss2, i, 4) = "" Or Gf_Get_Cell_Value(ss2, i, 4) = "0") _
               And (Gf_Get_Cell_Value(ss2, i, 2) = "0" And Gf_Get_Cell_Value(ss2, i, 3) = "0") Then
                .Row = i
                .RowHidden = True
            Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j
            End If
        Next i
                
    End With
    
End Sub

Private Sub subSpreadCheck1()
    
    Dim i As Long
    Dim j As Long
    
    With ss3
       
       For i = 1 To 166

           If Gf_Get_Cell_Value(ss3, i, 5) <> "A" And Gf_Get_Cell_Value(ss3, i, 5) <> "B" Then
               .Row = i
               .RowHidden = True
           Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j

           End If
'           以前用户提出 只有管线钢 需要这个项目显示 现在修改为 增加9Ni钢的显示 耿学玉  2011-5-10
            If Mid(Trim(txt_STDSPEC), 1, 3) <> "API" And Mid(Trim(txt_STDSPEC), 1, 10) <> "GB/T9711.2" And Trim(txt_STDSPEC) <> "70081MR-06Ni9" And Trim(txt_STDSPEC) <> "ASTM A553-9Ni" Then
                If i = 20 Or i = 21 Or i = 22 Or i = 23 Or i = 24 _
                   Or i = 25 Or i = 26 Or i = 70 Or i = 71 Or i = 72 _
                   Or i = 73 Or i = 74 Or i = 75 Or i = 76 Then
                   .RowHidden = True
                End If
            End If
'            增加 测膨胀值 只有9Ni钢才显示，因为用的都是冲击的判定代码 所以显示的内容较多
'            增加70021MR-06Ni9  20110819  liuxiang
            If Trim(txt_STDSPEC) <> "70081MR-06Ni9" And Trim(txt_STDSPEC) <> "70021MR-06Ni9" And Trim(txt_STDSPEC) <> "ASTM A553-9Ni" Then
                If i = 144 Or i = 145 Or i = 146 Or i = 147 Or i = 148 _
                   Or i = 149 Or i = 150 Or i = 151 Or i = 152 Or i = 153 _
                   Or i = 154 Or i = 155 Or i = 156 Or i = 157 Then
                   .RowHidden = True
                End If
            End If
            
       Next i
                
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





Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
  
    With ss4
        .Col = 1
        .Row = .ActiveRow
     txt_HEAT_NO.Text = .Text
    End With
    
    
    If ss4.MaxRows < 1 Or Row = 0 Or txt_HEAT_NO.Text = "" Then Exit Sub
    
    Call Gf_Sp_Refer(M_CN1, Sc1, Mc2, Mc2("pControl"))

End Sub

Private Sub txt_STDSPEC_Change()
    If Trim(txt_STDSPEC.Text) = "" Then
        txt_STDSPEC_NAME.Text = ""
    End If

End Sub
Private Sub subSpreadERROR(sPname As vaSpread)
    
    Dim i As Long
    Dim C_MAX, C_MIN, C_RESULT, C_FL As Variant

    With sPname
    
       If .MaxRows < 1 Then Exit Sub
       
       For i = 1 To .MaxRows
           .Row = i
           C_MIN = Val(Gf_Get_Cell_Value(sPname, i, 2))
           C_MAX = Val(Gf_Get_Cell_Value(sPname, i, 3))
           C_RESULT = Val(Gf_Get_Cell_Value(sPname, i, 4))
           If C_MIN <> 0 And C_MAX <> 0 Then
              If C_RESULT > C_MAX Or C_RESULT < C_MIN Then
                 Call Gp_Sp_CellColor(sPname, 4, i, RED)
              End If
           Else
              If C_MIN = 0 And C_MAX <> 0 Then
                 If C_RESULT > C_MAX Then
                    Call Gp_Sp_CellColor(sPname, 4, i, RED)
                 End If
              Else
                 If C_MIN <> 0 And C_MAX = 0 Then
                    If C_RESULT < C_MIN Then
                      Call Gp_Sp_CellColor(sPname, 4, i, RED)
                    End If
                 End If
              End If
           End If

       Next i
 
    End With
    
End Sub

Public Sub DataSave_H()
    Dim iRow, iCol As Integer

    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_H", Key:="P-M"

    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 7
           .Text = sUserID
    End With

    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet

    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub

    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0C0C0)                  ' 灰色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
         Next iRow
    End With
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_H", Key:="P-M"
    
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = ""
    End With
    

End Sub

Public Sub DataSave_C()
    Dim iRow, iCol As Integer

    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_C", Key:="P-M"

    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 7
           .Text = sUserID
    End With

    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet

    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub

    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H808080)                        ' 红色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
         Next iRow
    End With
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_C", Key:="P-M"
    
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = ""
    End With
    

End Sub

'材质确认取消
Public Sub DataSave_D()
    Dim iRow, iCol As Integer

    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_D", Key:="P-M"

    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 7
           .Text = sUserID
    End With

    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet

    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub

    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H808080)                        ' 红色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
         Next iRow
    End With
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_D", Key:="P-M"
    
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = ""
    End With
    

End Sub



