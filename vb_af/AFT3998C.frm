VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFT3998C 
   Caption         =   "供中板板卷板坯磅差计算界面_AFT3998C"
   ClientHeight    =   9225
   ClientLeft      =   495
   ClientTop       =   2520
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8655
      Left            =   60
      TabIndex        =   6
      Top             =   510
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   15266
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AFT3998C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   5565
         Left            =   0
         TabIndex        =   7
         Top             =   3090
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   9816
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AFT3998C.frx":0052
         Begin Threed.SSPanel SSPanel1 
            Height          =   555
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   15075
            _ExtentX        =   26591
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_RC 
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
               Left            =   5130
               TabIndex        =   13
               Tag             =   "工序代码"
               Top             =   120
               Width           =   990
            End
            Begin VB.TextBox txt_wgt 
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
               Left            =   8445
               TabIndex        =   12
               Tag             =   "过磅重量"
               Top             =   120
               Width           =   1005
            End
            Begin VB.TextBox txt_cal_wgt 
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
               Left            =   11160
               TabIndex        =   11
               Tag             =   "理重重量"
               Top             =   120
               Width           =   1005
            End
            Begin VB.TextBox txt_fp 
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
               Left            =   13890
               TabIndex        =   10
               Tag             =   "磅差率"
               Top             =   120
               Width           =   1005
            End
            Begin InDate.ULabel ULabel2 
               Height          =   315
               Left            =   3675
               Top             =   120
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               Caption         =   "抽查次数"
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
               Left            =   6990
               Top             =   120
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               Caption         =   "过磅重量"
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
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   9705
               Top             =   120
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               Caption         =   "理论重量"
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
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   12450
               Top             =   120
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               Caption         =   "磅差率"
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
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   225
               Top             =   120
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               Caption         =   "过磅日期"
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
            Begin InDate.UDate txt_OC_DATE 
               Height          =   315
               Left            =   1680
               TabIndex        =   14
               Tag             =   "日期"
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
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   4980
            Left            =   0
            TabIndex        =   15
            Top             =   585
            Width           =   15075
            _Version        =   393216
            _ExtentX        =   26591
            _ExtentY        =   8784
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
            MaxCols         =   13
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFT3998C.frx":00A4
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3030
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   15075
         _Version        =   393216
         _ExtentX        =   26591
         _ExtentY        =   5345
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
         MaxCols         =   10
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFT3998C.frx":088A
      End
   End
   Begin Threed.SSCommand cmd_update 
      Height          =   435
      Left            =   13410
      TabIndex        =   3
      Top             =   60
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "修改过磅重量"
   End
   Begin VB.TextBox txt_RC_id 
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
      Left            =   10890
      TabIndex        =   2
      Tag             =   "工序代码"
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txt_heat_no 
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
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "炉号"
      Top             =   120
      Width           =   1170
   End
   Begin InDate.ULabel ULabel63 
      Height          =   315
      Left            =   225
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      Caption         =   "炉号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3780
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      Caption         =   "查询过磅日期"
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
   Begin InDate.ULabel ULabel85 
      Height          =   315
      Left            =   9405
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      Caption         =   "抽查次数"
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
   Begin InDate.UDate SDT_FROM_DATE 
      Height          =   315
      Left            =   5250
      TabIndex        =   4
      Tag             =   "日期"
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
   Begin InDate.UDate SDT_TO_DATE 
      Height          =   315
      Left            =   7050
      TabIndex        =   5
      Tag             =   "日期"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6780
      TabIndex        =   1
      Top             =   180
      Width           =   255
   End
End
Attribute VB_Name = "AFT3998C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Shipping System
'-- Sub_System Name   Common
'-- Program Name      Insert shipping result
'-- Program ID        Refer
'-- Document No       Q-00-0010(Specification)
'-- Designer          zhang
'-- Coder             zhang
'-- Date              2009.9.19
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

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
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_heat_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
    
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
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
       'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFT3998C.P_REFER1", Key:="P-R"
    sc1.Add Item:="AFT3998C.P_ONEROW1", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
 'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(SDT_FROM_DATE, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(SDT_TO_DATE, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(txt_RC_id, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
  
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFT3998C.P_MODIFY2", Key:="P-M"
'    sc2.Add Item:="AFT3998C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:="AFT3998C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc"
    
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cmd_update_Click()

Dim i As Integer


If SDT_FROM_DATE.RawData <> SDT_TO_DATE.RawData Or txt_RC_id.Text = "" Then
 
       MsgBox "查询过镑日期必须相等和抽查次数不能为空", vbCritical, "系统提示信息"
       Exit Sub
End If
 
 
 If SDT_FROM_DATE.RawData <> "" And SDT_TO_DATE.RawData <> "" Or txt_RC_id.Text <> "" Then
      If ss2.Text <> "input" Or ss2.Text <> "delete" Then
          For i = 1 To ss2.MaxRows

              ss2.Col = 0
              ss2.Row = i
      
              ss2.Text = "Update"

        Next i

     End If
 
 
     Call Form_Pro
     
     If ss2.MaxRows > 0 Then

        For i = 1 To ss2.MaxRows

        txt_fp.Text = Round((Val(txt_cal_wgt.Text) - Val(txt_wgt.Text)) * 1000 / Val(txt_cal_wgt.Text), 3)

        Next i
     End If
    
End If
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "F-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Spl_SizeSet(SSSplitter1, "F-System.INI", Me.Name)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
     
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


Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        ss1.MaxRows = 0
        ss2.MaxRows = 0
        SDT_FROM_DATE.RawData = ""
        SDT_TO_DATE.RawData = ""
        txt_wgt.Text = ""
        txt_cal_wgt.Text = ""
        txt_fp.Text = ""
        txt_OC_DATE.RawData = ""
        txt_RC_id.Text = ""
        txt_RC.Text = ""
    End If
    
End Sub
Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call Gp_Sp_Excel(Me, sc2.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub


Public Sub Form_Ref()

    Dim iRow As Integer
    On Error Resume Next
    Dim temp_heat As String
    
    txt_wgt.Text = ""
    txt_cal_wgt.Text = ""
    txt_fp.Text = ""
    
 
If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
   
    
    If txt_heat_no.Text <> "" Then
    
     If Len(Trim(txt_heat_no.Text)) <> 8 Then
        Call Gp_MsgBoxDisplay("炉号长度应为8位", "", "错误提示")
     Else
       
        If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           ss1.OperationMode = OperationModeNormal
       
         End If
           
      End If
    End If
    
      
  If SDT_FROM_DATE.RawData <> "" And SDT_TO_DATE.RawData <> "" Then 'Or txt_RC_id.Text <> ""
  
         If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
       
              If ss2.MaxRows > 0 Then

             For iRow = 1 To ss2.MaxRows

                 ss2.Row = iRow
                 ss2.Col = 10
                 txt_wgt.Text = Val(txt_wgt.Text) + Val(ss2.Text)

                 ss2.Col = 11
                 txt_cal_wgt.Text = Val(txt_cal_wgt.Text) + Val(ss2.Text)
                    
                 txt_fp.Text = Round((Val(txt_cal_wgt.Text) - Val(txt_wgt.Text)) * 1000 / Val(txt_cal_wgt.Text), 3)

               Next iRow
           End If
       
     End If
     
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
      ss2.OperationMode = OperationModeNormal
 
             
End If

End Sub

Public Sub Form_Pro()

    Dim TEMP As Double
    Dim TEMPs As Double
    Dim i As Integer
    Dim sErrMessg As String


    If ss2.MaxRows = 0 Then
       MsgBox "末选定过磅板坯", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    If txt_wgt.Text = "" Then
       MsgBox "过镑重量不能为空", vbCritical, "系统提示信息"
       Exit Sub
    End If
       
    If ss2.MaxRows = 1 Then

        ss2.Row = ss2.ActiveRow
        ss2.Col = 10
        ss2.Text = txt_wgt.Text

      ElseIf ss2.MaxRows > 0 Then

         For i = 1 To ss2.MaxRows
            If i < ss2.MaxRows Then
                ss2.Row = i
                ss2.Col = 10
                ss2.Text = Round(Val(txt_wgt.Text) / ss2.MaxRows, 2)
                TEMP = TEMP + Val(ss2.Text)

              
            ElseIf i = ss2.MaxRows Then
                ss2.Row = i
                ss2.Col = 10
                ss2.Text = Val(txt_wgt.Text) - TEMP
                

            End If
         Next i
   End If
      

      
      If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc2) Then
   
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
   
      ss2.OperationMode = OperationModeNormal
      
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

Private Sub opt_KND_Click(Index As Integer)
   
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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub
Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Integer
Dim sslab_no As String
    txt_cal_wgt.Text = ""
    txt_wgt.Text = ""

    If Row < 1 Then Exit Sub

    If txt_OC_DATE.RawData = "" Or txt_RC.Text = "" Then
       MsgBox "过镑日期和抽查次数不能为空", vbCritical, "系统提示信息"
       Exit Sub
    End If

    ss1.Row = Row
    ss1.Col = 2
    sslab_no = Trim(ss1.Text)
    ss1.Col = 1
    If ss1.VALUE = 1 Then
       If ss2.MaxRows > 0 Then
            For i = 1 To ss2.MaxRows
                 ss2.Row = i
                 ss2.Col = 3
                 If ss2.Text = sslab_no Then
                    MsgBox "该板坯号已选定", vbCritical, "系统提示信息"
                    Exit Sub
                 End If
            Next i
       End If
       
     
      Call Gp_Sp_Ins(Proc_Sc("Sc"))
      
      ss2.Col = 12
      ss2.Row = ss2.ActiveRow
      ss2.Text = sUserID
    
      ss2.Row = ss2.ActiveRow
      ss2.Col = 1
      ss2.VALUE = txt_OC_DATE.RawData
      ss2.Col = 2
      ss2.Text = txt_RC.Text
      
     ss1.Row = ss1.ActiveRow
     ss1.Col = 2
     ss2.Col = 3
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     ss1.Row = ss1.ActiveRow
     ss1.Col = 3
     ss2.Col = 4
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     ss1.Row = ss1.ActiveRow
     ss1.Col = 4
     ss2.Col = 5
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     ss1.Row = ss1.ActiveRow
     ss1.Col = 6
     ss2.Col = 6
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     ss1.Row = ss1.ActiveRow
     ss1.Col = 7
     ss2.Col = 7
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     ss1.Row = ss1.ActiveRow
     ss1.Col = 8
     ss2.Col = 8
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     ss1.Row = ss1.ActiveRow
     ss1.Col = 9
     ss2.Col = 9
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     ss1.Row = ss1.ActiveRow
     ss1.Col = 10
     ss2.Col = 11
     ss2.Row = ss2.ActiveRow
     ss2.Text = ss1.Text
     
     
     
    Else
         For i = 1 To ss2.MaxRows

   
                 ss2.Row = i
                 ss2.Col = 3
                 If ss2.Text = sslab_no Then
                    Call ss2.SetActiveCell(1, i)

                    ss2.Col = ss2.ActiveCol: ss2.Col2 = ss2.ActiveCol
                    ss2.Row = i: ss2.Row2 = i
                    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
                    Exit For
                 End If
         Next i
         
    End If
          For i = 1 To ss2.MaxRows

                 ss2.Row = i
                 ss2.Col = 11
                 txt_cal_wgt.Text = Val(txt_cal_wgt.Text) + Val(ss2.Text)
                    
                
               Next i


End Sub


    
Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = Col
    lBlkcol2 = Col
    lBlkrow1 = Row
    lBlkrow2 = Row

End Sub
Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

  
     If Col = 1 Then
        ss2.Col = Col
        ss2.Row = Row
        ss2.Text = Format(Now, "YYYY-MM-DD")
         End If

End Sub
    
Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim str1 As String
   
    If Gf_Sc_Authority(sAuthority, "U") Then
       Call Gp_Sp_UpdateMake(ss2, Mode)
    End If
    
End Sub

Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
   
Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
      
    ss2.Row = Row: ss1.Col = Col
    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub SDT_FROM_DATE_DblClick()
     SDT_FROM_DATE.RawData = Gf_DTSet(M_CN1, "D")
     SDT_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub
Private Sub SDT_TO_DATE_DblClick()
     SDT_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub





Private Sub txt_OC_DATE_DblClick()
    txt_OC_DATE.RawData = Gf_DTSet(M_CN1, "D")
    
End Sub
Public Sub Form_Ins()

    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss2.Col = 12
    ss2.Row = ss2.ActiveRow
    ss2.Text = sUserID

End Sub
Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    ss2.Col = 12
    ss2.Row = ss2.ActiveRow
    ss2.Text = sUserID
    
End Sub

Public Sub Spread_Del()
   Call Gp_Sp_Del(Proc_Sc("Sc"))
End Sub

