VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1031C 
   Caption         =   "销售计划录入_AAA1031C"
   ClientHeight    =   8130
   ClientLeft      =   540
   ClientTop       =   2190
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   13995
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand SCmd2 
      Height          =   375
      Left            =   12330
      TabIndex        =   29
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "炼钢能力检查"
   End
   Begin Threed.SSCommand SCmd1 
      Height          =   375
      Left            =   13770
      TabIndex        =   28
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "计划汇总"
   End
   Begin FPSpread.vaSpread SS2 
      Height          =   1575
      Left            =   5670
      TabIndex        =   27
      Top             =   1410
      Width           =   9375
      _Version        =   393216
      _ExtentX        =   16536
      _ExtentY        =   2778
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
      MaxCols         =   6
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      SpreadDesigner  =   "AAA1031C.frx":0000
   End
   Begin VB.TextBox txt_plate_ex 
      Height          =   270
      Left            =   12735
      TabIndex        =   26
      Top             =   855
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt_coil_ex 
      Height          =   270
      Left            =   11655
      TabIndex        =   25
      Top             =   855
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt_plate 
      Height          =   270
      Left            =   10485
      TabIndex        =   24
      Top             =   855
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt_coil 
      Height          =   270
      Left            =   12645
      TabIndex        =   23
      Top             =   540
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt_th_plate 
      Height          =   270
      Left            =   11610
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt_th_coil 
      Height          =   270
      Left            =   10485
      TabIndex        =   21
      Top             =   540
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt_tot_hour 
      Height          =   270
      Left            =   9495
      TabIndex        =   20
      Top             =   540
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt_prod_cd 
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
      Left            =   5985
      MaxLength       =   2
      TabIndex        =   11
      Tag             =   "产品"
      Top             =   135
      Width           =   555
   End
   Begin VB.TextBox txt_cust_name 
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
      Left            =   2610
      MaxLength       =   40
      TabIndex        =   2
      Top             =   495
      Width           =   6765
   End
   Begin VB.TextBox txt_cust_cd 
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
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   1
      Top             =   495
      Width           =   1050
   End
   Begin InDate.ULabel ULabel4 
      Height          =   330
      Left            =   2835
      Top             =   1395
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "已使用(h)"
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
      Height          =   330
      Left            =   270
      Top             =   1395
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "总时间(h)"
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
   Begin VB.TextBox txt_stlgrd 
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
      Left            =   1530
      MaxLength       =   11
      TabIndex        =   3
      Top             =   855
      Width           =   1410
   End
   Begin VB.TextBox txt_stlgrd_des 
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
      Left            =   2970
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   855
      Width           =   6405
   End
   Begin InDate.ULabel ULabel6 
      Height          =   285
      Left            =   180
      Top             =   855
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      Caption         =   "钢种"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel5 
      Height          =   300
      Left            =   4635
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "产品"
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
      ForeColor       =   16711680
   End
   Begin InDate.UDate dtp_date_str 
      Height          =   300
      Left            =   1530
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   135
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Left            =   180
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "日期"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   5955
      Left            =   135
      TabIndex        =   5
      Top             =   3195
      Width           =   15090
      _Version        =   393216
      _ExtentX        =   26617
      _ExtentY        =   10504
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
      MaxCols         =   0
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1031C.frx":0530
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   180
      Top             =   495
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "客户"
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
      ForeColor       =   16711680
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   135
      X2              =   15165
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   135
      X2              =   15165
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label lbl_use_hour 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4140
      TabIndex        =   19
      Top             =   1395
      Width           =   1005
   End
   Begin VB.Label lbl_tot_hour 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1575
      TabIndex        =   18
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Label lbl_plate_3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4095
      TabIndex        =   17
      Top             =   2655
      Width           =   960
   End
   Begin VB.Label lbl_plate_2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   16
      Top             =   2655
      Width           =   1140
   End
   Begin VB.Label lbl_plate_1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1605
      TabIndex        =   15
      Top             =   2655
      Width           =   1065
   End
   Begin VB.Label lbl_coil_3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4095
      TabIndex        =   14
      Top             =   2250
      Width           =   960
   End
   Begin VB.Label lbl_coil_2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   13
      Top             =   2250
      Width           =   1170
   End
   Begin VB.Label lbl_coil_1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1605
      TabIndex        =   12
      Top             =   2250
      Width           =   1065
   End
   Begin VB.Line Line5 
      X1              =   4140
      X2              =   5130
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line4 
      X1              =   1575
      X2              =   2655
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "时间(h)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4095
      TabIndex        =   10
      Top             =   1890
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "计划(t)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2970
      TabIndex        =   9
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "比例"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1710
      TabIndex        =   8
      Top             =   1890
      Width           =   780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      Index           =   2
      X1              =   4005
      X2              =   4005
      Y1              =   1800
      Y2              =   2970
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      Index           =   1
      X1              =   2700
      X2              =   2700
      Y1              =   1800
      Y2              =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "钢板"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   315
      TabIndex        =   7
      Top             =   2700
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "钢卷"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   315
      TabIndex        =   6
      Top             =   2250
      Width           =   1185
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      X1              =   270
      X2              =   5130
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1530
      X2              =   1530
      Y1              =   1800
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   270
      X2              =   5130
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1185
      Left            =   270
      Top             =   1800
      Width           =   4875
   End
End
Attribute VB_Name = "AAA1031C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       production plan
'-- Sub_System Name
'-- Program Name
'-- Program ID        AAA1030C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2003.7.9
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim THK_GRP As Collection
Dim WID_GRP As Collection
Dim MIN_VALUE As Collection
Dim MAX_VALUE As Collection

Dim arrValue As Variant

Private Sub Form_Define()
    
    Dim sQuery As String
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(dtp_date_str, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_cust_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_stlgrd_des, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Spread_Collection
    Sc1.Add Item:="AAA1030C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    'sQuery = "SELECT cd from zp_cd where cd_mana_no='B0005'"
    'Call Gf_ComboAdd(M_CN1, cbo_prod_cd, sQuery)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub SCmd2_Click()

On Error GoTo SCmd2_Error

    Dim sQuery, sEdate, sPrc As String
    Dim iCount As Integer
    Dim sTemp As Double
    
    If dtp_date_str.Enabled Then Exit Sub
    
    Dim adoCmd As adodb.Command
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "AAA1031P"
    
    'CCM
    sEdate = Mid(dtp_date_str.Text, 1, 4) + Mid(dtp_date_str.Text, 6, 2)
    sPrc = "BF"
    sTemp = 0#
    
    'Ceate Parameter (Input) iType + iColumn
    For iCount = 1 To 2
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("v1", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("v2", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("v3", adVariant, adParamOutput)
    
    adoCmd.Parameters(0).Value = sPrc
    adoCmd.Parameters(1).Value = sEdate
    
    adoCmd.Execute , , adExecuteNoRecords
    
    With ss2
        .Row = 1
        .Col = 1
        .Text = adoCmd("v1")
        .Col = 3
        .Text = adoCmd("v2")
        .Col = 5
        .Text = adoCmd("v3")
        .Col = 2
        sTemp = CLng(lbl_coil_2.Caption) + CLng(lbl_plate_2.Caption)
        .Text = Format(sTemp)
        .Col = 4
        If Val(adoCmd("v2")) <> 0 Then
          .Text = FormatNumber((CLng(lbl_coil_2.Caption) + CLng(lbl_plate_2.Caption)) / (Val(adoCmd("v2")) / 100), 0)
          sTemp = (CLng(lbl_coil_2.Caption) + CLng(lbl_plate_2.Caption)) / (Val(adoCmd("v2") / 100))
        Else
          sTemp = 0
        End If
        .Col = 6
        If Val(adoCmd("v2")) <> 0 And Val(adoCmd("v3")) <> 0 Then
          .Text = FormatNumber(((CLng(lbl_coil_2.Caption) + CLng(lbl_plate_2.Caption)) / (Val(adoCmd("v2")) / 100)) / Val(adoCmd("v3")), 2)
        End If
    End With
    
    'BOF
    sPrc = "BC"
    adoCmd.Parameters(0).Value = sPrc
    adoCmd.Parameters(1).Value = sEdate
    adoCmd.Execute , , adExecuteNoRecords
    
    With ss2
        .Row = 2
        .Col = 1
        .Text = adoCmd("v1")
        .Col = 3
        .Text = adoCmd("v2")
        .Col = 5
        .Text = adoCmd("v3")
        .Col = 2
        .Text = FormatNumber(sTemp, 0)
        .Col = 4
        If Val(adoCmd("v2")) <> 0 Then
           .Text = FormatNumber(sTemp / (Val(adoCmd("v2")) / 100), 0)
           'sTemp = IIf(.Text = "", 0, .Text)
        Else
          sTemp = 0
        End If
        .Col = 6
        If sTemp <> 0 And Val(adoCmd("v2")) <> 0 And Val(adoCmd("v3")) <> 0 Then
            .Text = FormatNumber((sTemp / (Val(adoCmd("v2")) / 100)) / Val(adoCmd("v3")), 2)
        End If
    End With
    
    Set adoCmd = Nothing
    
    If subCollectionAdd = True Then
        Call subValueCheck
    End If
    
    Exit Sub

SCmd2_Error:

    Call Gp_MsgBoxDisplay("SCmd2_Error : " & Error)

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Menu_Setting

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
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_Setting(ss2)
    
    ss2.MaxRows = 2
    
    ss2.Col = 0
    ss2.Row = 1
    ss2.Text = "CAST"
    ss2.Row = 2
    ss2.Text = "BOF"
    
    Call Gp_Sp_EvenRowBackcolor(ss2, 0)
    ss2.OperationMode = OperationModeRead
    
    Call Sp_Setting
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(ss2, "A-System.INI", Me.Name)

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(ss2, "A-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set THK_GRP = Nothing
    Set WID_GRP = Nothing
    Set MIN_VALUE = Nothing
    Set MAX_VALUE = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    rControl(1).SetFocus
    lbl_tot_hour.Caption = ""
    lbl_use_hour.Caption = ""
    lbl_coil_1.Caption = ""
    lbl_coil_2.Caption = ""
    lbl_coil_3.Caption = ""
    lbl_plate_1.Caption = ""
    lbl_plate_2.Caption = ""
    lbl_plate_3.Caption = ""
    
    ss1.MaxCols = 0
    ss1.MaxRows = 0
    
    ss2.ClearRange 1, 1, ss2.MaxCols, ss2.MaxRows, False
 
End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
        
        If Sp_Header_Refer() Then
            If Sp_Data_Refer() Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Menu_Setting
                Call Gp_Ms_ControlLock(Mc1!lControl, True)
            End If
        End If
            
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Public Sub Form_Pro()

    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Form_Ref
    End If
    
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim dTotal As Double
    Dim DCURR As Double
    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate As String
    Dim AdoRs1 As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs1 = New adodb.Recordset
    
    sEdate = Mid(dtp_date_str.Text, 1, 4) + Mid(dtp_date_str.Text, 6, 2)
             
    dTotal = 0
    With ss1
        For iRow = 1 To .MaxRows
            For iCol = 2 To .MaxCols Step 2
                .Row = iRow
                .Col = iCol
                If .Value = "" Then
                    DCURR = 0
                Else
                    DCURR = .Value
                End If
                dTotal = dTotal + DCURR
             Next iCol
        Next iRow
            
        Select Case Trim(txt_prod_cd.Text)
            Case "HC"
                If Val(txt_th_coil.Text) <> 0 And Val(txt_th_plate.Text) <> 0 Then
                
                    If Val(txt_th_coil.Text) + Val(txt_plate.Text) = 0 Or Val(txt_th_plate.Text) = 0 Then
                        lbl_use_hour = 0
                    Else
                        lbl_use_hour = Int((dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text)) / Val(txt_th_coil.Text) + Val(txt_plate.Text) / Val(txt_th_plate.Text))
                    End If
                    
                    If (dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text)) = 0 Or (dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text) + Val(txt_plate.Text)) = 0 Then
                        lbl_coil_1.Caption = 0
                    Else
                        lbl_coil_1.Caption = FormatPercent((dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text)) / (dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text) + Val(txt_plate.Text)), 2)
                    End If
                    
                    lbl_coil_2.Caption = FormatNumber(dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text), 0)
                    
                    If (dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text)) = 0 Or Val(txt_th_coil.Text) = 0 Then
                        lbl_coil_3.Caption = 0
                    Else
                        lbl_coil_3.Caption = FormatNumber((dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text)) / Val(txt_th_coil.Text), 2)
                    End If
                    
                    If (Val(txt_plate.Text)) = 0 Or (dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text) + Val(txt_plate.Text)) = 0 Then
                        lbl_plate_1.Caption = 0
                    Else
                        lbl_plate_1.Caption = FormatPercent((Val(txt_plate.Text)) / (dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text) + Val(txt_plate.Text)), 2)
                    End If
                   
                End If
                     
            Case "PP"
                If Val(txt_th_coil.Text) <> 0 And Val(txt_th_plate.Text) <> 0 Then
                
                    If (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text)) = 0 Or Val(txt_th_plate.Text) + Val(txt_coil.Text) = 0 Or Val(txt_th_coil.Text) = 0 Then
                        lbl_use_hour = 0
                    Else
                        lbl_use_hour = Int((dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text)) / Val(txt_th_plate.Text) + Val(txt_coil.Text) / Val(txt_th_coil.Text))
                    End If
                    
                    If (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text)) = 0 Or (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text) + Val(txt_coil.Text)) = 0 Then
                        lbl_plate_1.Caption = 0
                    Else
                        lbl_plate_1.Caption = FormatPercent((dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text)) / (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text) + Val(txt_coil.Text)), 2)
                    End If
                    
                    lbl_plate_2.Caption = FormatNumber(dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text), 0)
                    
                    If (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text)) = 0 Or Val(txt_th_plate.Text) = 0 Then
                        lbl_plate_3.Caption = 0
                    Else
                        lbl_plate_3.Caption = FormatNumber((dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text)) / Val(txt_th_plate.Text), 2)
                    End If
                    
                    If (Val(txt_coil.Text)) = 0 Or (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text) + Val(txt_coil.Text)) = 0 Then
                    
                    Else
                        lbl_coil_1.Caption = FormatPercent((Val(txt_coil.Text)) / (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text) + Val(txt_coil.Text)), 2)
                    End If
                    
                 End If
               
        End Select
        
        Select Case Trim(txt_prod_cd.Text)
            Case "HC"
                If Val(txt_th_coil.Text) <> 0 And Val(txt_th_plate.Text) <> 0 Then
                    If txt_tot_hour < (dTotal + Val(txt_coil.Text) - Val(txt_coil_ex.Text)) / Val(txt_th_coil.Text) + Val(txt_plate.Text) / Val(txt_th_plate.Text) Then
                        .Col = Col
                        .Row = Row
                        .CellTag = "False"
                        
                        Call Gp_MsgBoxDisplay("需要的时间超过总时间...")
                        
                        .Col = Col
                        .Row = Row
                        .CellTag = ""
                        
                        .Value = 0
                        .TabStop = True
                        .SetFocus
                        .SetActiveCell Col, Row
                        .Action = SS_ACTION_ACTIVE_CELL
                        .EditMode = True
                        .TabStop = False
                
                    End If
                End If
            Case "PP"
                If Val(txt_th_coil.Text) <> 0 And Val(txt_th_plate.Text) <> 0 Then
                    If txt_tot_hour < (dTotal + Val(txt_plate.Text) - Val(txt_plate_ex.Text)) / Val(txt_th_plate.Text) + Val(txt_coil.Text) / Val(txt_th_coil.Text) Then
                        .Col = Col
                        .Row = Row
                        .CellTag = "False"
                    
                        Call Gp_MsgBoxDisplay("需要的时间超过总时间...")
                    
                        .Col = Col
                        .Row = Row
                        .CellTag = ""
                    
                        .Value = 0
                        .TabStop = True
                        .SetFocus
                        .SetActiveCell Col, Row
                        .Action = SS_ACTION_ACTIVE_CELL
                        .EditMode = True
                        .TabStop = False
                    
                    End If
                End If
        End Select
                    
'                    If dMax <> 0 Then
'                        If dMax < dMin Then
'
'                            .Col = Col
'                            .Row = Row
'                            .CellTag = "False"
'
'                            Call Gp_MsgBoxDisplay("最大值应大于最小值...")
'
'                            .Col = Col
'                            .Row = Row
'                            .CellTag = ""
'
'                            .Value = 0
'                            .TabStop = True
'                            .SetFocus
'                            .SetActiveCell Col, Row
'                            .Action = SS_ACTION_ACTIVE_CELL
'                            .EditMode = True
'                            .TabStop = False
'
'                        End If
'                    End If
    End With

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
        MDIMain.Mnu_Sorting.Visible = False
        MDIMain.Line1.Visible = False
        
        PopupMenu MDIMain.PopUp_Spread
        
        MDIMain.Mnu_Sorting.Visible = True
        MDIMain.Line1.Visible = True
    End If

End Sub

Public Sub Sp_Setting()

    With ss1

        .ColHeaderRows = 3
        .RowHeaderCols = 2
        .Col = -1
        .Row = SpreadHeader + 1
        .FontBold = True
        
        .RowHeight(SpreadHeader) = 15
        .RowHeight(SpreadHeader + 1) = 15
        
        .Row = SpreadHeader + 2
     ''   .RowHidden = True
        
        .ColWidth(0) = 10
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
          
        .Row = SpreadHeader
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        .Row = SpreadHeader + 1
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        
        .Row = SpreadHeader + 2
        .RowHidden = True
        
        .Col = SpreadHeader + 1
        .ColHidden = True
        
    End With
    

End Sub

Public Sub Menu_Setting()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel
    
End Sub

Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    
    Dim sQuery2 As String
    
    Dim AdoRs2 As adodb.Recordset
    Dim ArrayRecords2 As Variant

    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT THK_CD, FR_THK, TO_THK "
    sQuery = sQuery + "   FROM BP_THICK_GRP "
    sQuery = sQuery + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery = sQuery + "    AND THK_CD <> '*' "
    sQuery = sQuery + "  ORDER BY THK_CD "
    
    With ss1

        Sp_Header_Refer = True
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 2
        
            For iCol = 0 To .MaxCols - 1 Step 2
            
               .Col = iCol + 1
               .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                  
                .Col = iCol + 2
                .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                           
                .Col = iCol + 1:  .Row = SpreadHeader + 1:  .Text = "Actual"
                .Col = iCol + 2:  .Row = SpreadHeader + 1:  .Text = "Plan"
                
                .Col = iCol + 1
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                .Col = iCol + 2
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                'Column Type Setting
                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                .ColWidth(iCol + 1) = 9
                
                .Col = iCol + 2: .Col2 = iCol + 2
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                .ColWidth(iCol + 2) = 9
                
                iCnt = iCnt + 1
                
            Next iCol
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Set AdoRs2 = New adodb.Recordset
    
    sQuery2 = "SELECT WID_CD, FR_WID, TO_WID "
    sQuery2 = sQuery2 + "   FROM BP_WIDTH_GRP "
    sQuery2 = sQuery2 + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery2 = sQuery2 + "    AND WID_CD <> '*' "
    sQuery2 = sQuery2 + "  ORDER BY WID_CD "
    
    With ss1

        Sp_Header_Refer = True
     '   .ReDraw = False
     '   .MaxRows = 0:  .MaxCols = 0
         .ColWidth(0) = 15
      '  .ColWidth(1) = 20
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs2.Open sQuery2, M_CN1, adOpenKeyset
        
        If AdoRs2.BOF Or AdoRs2.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            AdoRs2.Close
            Set AdoRs2 = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords2 = AdoRs2.GetRows
        AdoRs2.Close
        Set AdoRs2 = Nothing

        If UBound(ArrayRecords2, 2) + 1 <> 0 Then
        
            .MaxRows = (UBound(ArrayRecords2, 2) + 1)
            iCnt = 0
            
            For iRow = 1 To .MaxRows
            
                .Row = iRow
                .Col = SpreadHeader
                
                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords2(1, iCnt)) & " ~ " & Trim(ArrayRecords2(2, iCnt)) & "mm"
                End If
                
                .Col = SpreadHeader + 1
                .Text = Trim(ArrayRecords2(0, iCnt))
                
                .Row = iRow + 2: .Row2 = iRow + 2
                .Col = 1: .Col2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                iCnt = iCnt + 1
            Next iRow
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    With ss1
    
        For iCol = 1 To .MaxCols - 1 Step 2
            .Col = iCol
            .Row = 1
            .Col2 = iCol
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            .Protect = True
        Next iCol
        
        For iCol = 2 To .MaxCols Step 2
            .Col = iCol
            .Row = 1
            .Col2 = iCol
            .Row2 = .MaxRows
             If Trim(txt_prod_cd.Text) = "" Or Trim(txt_cust_cd.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then
                .BlockMode = True
                .Lock = True
                .BlockMode = False
                .Protect = True
             Else
                .BlockMode = True
                .Lock = False
                .BackColor = &HC0FFFF
                .BlockMode = False
                .Protect = True
             End If
        Next iCol
        
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Set AdoRs2 = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim dTotal As Double
    Dim DCURR As Double
    Dim sTdate As String
    Dim sQuery, sQuery1 As String
    Dim sEdate As String
    Dim sWID_GRP As String
    Dim sTHK_GRP As String
   ' Dim SPARA As String
    Dim AdoRs As adodb.Recordset
    Dim AdoRs1 As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    sEdate = Mid(dtp_date_str.Text, 1, 4) + Mid(dtp_date_str.Text, 6, 2)
  
    sQuery = "SELECT WID_GRP, THK_GRP, sum(RST_WGT),sum(PLN_WGT)"
    sQuery = sQuery + "   FROM AP_SALES_PLAN "
    sQuery = sQuery + "  WHERE YEAR_MONTH =      '" + sEdate + "' "
    sQuery = sQuery + "    AND CUST_CD    LIKE   '" + Trim(txt_cust_cd.Text) + "%' "
    sQuery = sQuery + "    AND PROD_CD    LIKE   '" + Trim(txt_prod_cd.Text) + "%' "
    sQuery = sQuery + "    AND STLGRD     LIKE   '" + Trim(txt_stlgrd.Text) + "%' "
    sQuery = sQuery + "  GROUP BY WID_GRP, THK_GRP "
    sQuery = sQuery + "  ORDER BY WID_GRP, THK_GRP "
    
    Set AdoRs1 = New adodb.Recordset
    sQuery1 = "{ call AAA1030C.P_REFER('" + sEdate + "') }"
    AdoRs1.Open sQuery1, M_CN1, adOpenKeyset
        
    If AdoRs1.BOF Or AdoRs1.EOF Then
    
        Sp_Data_Refer = False
        ss1.ReDraw = True
        AdoRs1.Close
        Set AdoRs1 = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If
        
    ArrayRecords = AdoRs1.GetRows
    
    AdoRs1.Close
    Set AdoRs1 = Nothing
    
    lbl_tot_hour.Caption = Str(ArrayRecords(0, 0))
    
    If ArrayRecords(1, 0) <> 0 And ArrayRecords(2, 0) <> 0 Then
        lbl_use_hour = Str(Int((ArrayRecords(3, 0) / ArrayRecords(1, 0)) + (ArrayRecords(4, 0) / ArrayRecords(2, 0))))
    End If
    
    If ArrayRecords(3, 0) + ArrayRecords(4, 0) <> 0 Then
        lbl_coil_1.Caption = FormatPercent(ArrayRecords(3, 0) / (ArrayRecords(3, 0) + ArrayRecords(4, 0)), 2)
    End If
    
    lbl_coil_2.Caption = FormatNumber(ArrayRecords(3, 0), 0)
    
    If ArrayRecords(1, 0) <> 0 Then
        lbl_coil_3.Caption = FormatNumber(ArrayRecords(3, 0) / ArrayRecords(1, 0), 2)
    End If
    
    If ArrayRecords(3, 0) + ArrayRecords(4, 0) <> 0 Then
        lbl_plate_1.Caption = FormatPercent(ArrayRecords(4, 0) / (ArrayRecords(3, 0) + ArrayRecords(4, 0)), 2)
    End If
    
    lbl_plate_2.Caption = FormatNumber(ArrayRecords(4, 0), 0)
    
    If ArrayRecords(2, 0) <> 0 Then
        lbl_plate_3.Caption = FormatNumber(ArrayRecords(4, 0) / ArrayRecords(2, 0), 2)
    End If
    
    txt_tot_hour.Text = ArrayRecords(0, 0)
    txt_th_coil.Text = ArrayRecords(1, 0)
    txt_th_plate.Text = ArrayRecords(2, 0)
    txt_coil.Text = ArrayRecords(3, 0)
    txt_plate.Text = ArrayRecords(4, 0)
        
    With ss1

        Sp_Data_Refer = True
        
        .ReDraw = False
       ' .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
     '   Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
            iRow = 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .Row = iRow
                .Col = SpreadHeader + 1
                 sWID_GRP = .Text
                 Do While iRow <= .MaxRows And sWID_GRP <> Trim(ArrayRecords(0, iCnt))
                    iRow = iRow + 1
                    .Row = iRow
                    sWID_GRP = .Text
                 Loop
                           
                 For iCol = 1 To .MaxCols - 1 Step 2
                    .Col = iCol
                    .Row = SpreadHeader + 2
                    sTHK_GRP = .Text

                    If sTHK_GRP = ArrayRecords(1, iCnt) Then
                        
                        .Row = iRow
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Or ArrayRecords(2, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(2, iCnt))
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Or ArrayRecords(3, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(3, iCnt))
                        End If
                
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
     '   .ReDraw = True
        dTotal = 0
        With ss1
            For iRow = 1 To .MaxRows
                For iCol = 2 To .MaxCols Step 2
                
                     .Row = iRow
                     .Col = iCol
                    If .Value = "" Then
                        DCURR = 0
                    Else
                        DCURR = .Value
                    End If
                    
                    dTotal = dTotal + DCURR
                Next iCol
            Next iRow
        End With
         
        If Trim(txt_prod_cd.Text) = "HC" Then txt_coil_ex.Text = dTotal
        If Trim(txt_prod_cd.Text) = "PP" Then txt_plate_ex.Text = dTotal
        
        MDIMain.StatusBar1.Panels(1) = "Message : Data inquiry completed"
        Screen.MousePointer = vbDefault
        
    End With
         
    Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Process(Conn As adodb.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    
    Dim adoCmd As adodb.Command

    Sp_Process = True
    
    If Trim(txt_prod_cd.Text) = "" Or Trim(txt_cust_cd.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then
       Sp_Process = False
       Call Gp_MsgBoxDisplay("can't save ...")
       Exit Function
    End If
    
    With ss1
    
        'MaxRow = 0 is Exit Function Or iCount = 0
        If .MaxRows < 1 Then
            Sp_Process = False
            Exit Function
        End If
        
        Screen.MousePointer = vbHourglass
        
        .ReDraw = False
        
        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Sp_Process = False: Exit Function
        End If
        
        'Ado Setting
        Conn.CursorLocation = adUseServer
        Set adoCmd = New adodb.Command
        
        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = Sc.Item("P-M")
        
        Conn.BeginTrans
        
        'Ceate Parameter (Input) iType + iColumn
        For iCount = 1 To 8
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            'Parameters Setting
            For iCol = 2 To .MaxCols Step 2
            
                .Col = iCol
                If Trim(.Text) <> "" Then
                
                    .Row = SpreadHeader + 2
                    .Col = iCol
                    adoCmd.Parameters(4).Value = .Text     'thk_grp
              
                    .Row = iRow
                    .Col = SpreadHeader + 1
                    adoCmd.Parameters(5).Value = .Text     'wid_grp
                    
                    .Col = iCol
                 
                    If Trim(.Text) = "" Then               'plan_value
                        adoCmd.Parameters(6).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(6).Value = dTempInt
                    End If
                    
                    adoCmd.Parameters(7).Value = sUserID                            'User-id
                    
                    adoCmd.Parameters(0).Value = Mid(dtp_date_str.Text, 1, 4) + _
                                                 Mid(dtp_date_str.Text, 6, 2)        'YEAR_MONTH
                    adoCmd.Parameters(1).Value = txt_cust_cd.Text                    'CUST_CD
                    adoCmd.Parameters(2).Value = txt_prod_cd.Text                    'PROD_CD
                    adoCmd.Parameters(3).Value = txt_stlgrd.Text                     'STLGRD
                                   
                    adoCmd.Execute
                    
                    'Error Check
                    If adoCmd("Error") <> "0" Then
               
                        ret_Result_ErrCode = adoCmd("Error")
                        ret_Result_ErrMsg = adoCmd("Messg")
                        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
               
                        Call Gp_MsgBoxDisplay(sErrMessg)
                        Screen.MousePointer = vbDefault
                        Set adoCmd = Nothing
                        Conn.RollbackTrans
                        Sp_Process = False
                        Exit Function
               
                     End If
                
                End If
            
            Next iCol
            
        Next iRow
        
        Conn.CommitTrans
        .ReDraw = True
        MDIMain.StatusBar1.Panels(1) = "Message : Data update completed"
        Screen.MousePointer = vbDefault
        Exit Function
    
    End With

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("SpreadPro_Error : " & Error)

End Function

Private Sub SCmd1_Click()

    Load AAA1080C
    AAA1080C.txt_year_month.Text = Mid(dtp_date_str, 1, 4) + Mid(dtp_date_str, 6, 2)
    AAA1080C.Show 1
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_name

        DD.nameType = "2"
        Call Gf_Customer_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 2)
    Else
        txt_cust_name.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_des

        DD.nameType = "2"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
        txt_stlgrd_des.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
    Else
        txt_stlgrd_des.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Function subCollectionAdd() As Boolean

    Dim i As Integer
    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    
    If dtp_date_str.RawData = "" Or Trim(txt_prod_cd.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then Exit Function

    Set AdoRs = New adodb.Recordset

'Collection Clear ----------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
    sQuery = "SELECT THK_GRP , WID_GRP , SUM(NVL(MIN,0)) , SUM(NVL(MAX,0)) FROM AP_LIMIT_CON WHERE "
    sQuery = sQuery + "YEAR_MONTH = '" + Left(dtp_date_str.RawData, 6) + "' AND "
    sQuery = sQuery + "PROD_CD    = '" + txt_prod_cd.Text + "' AND "
    sQuery = sQuery + "STLGRD     = '" + txt_stlgrd.Text + "' "
    sQuery = sQuery + "GROUP BY THK_GRP, WID_GRP "
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.BOF Or AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        Exit Function
    End If
    
    arrValue = AdoRs.GetRows
    AdoRs.Close
    Set AdoRs = Nothing
    subCollectionAdd = True
                
End Function

Private Sub subValueCheck()

    Dim i As Integer
    Dim j As Integer
    Dim bCheck As Boolean
    Dim iCnt As Integer
    Dim sMesg As String
    
    With ss1
        
        For i = 1 To .MaxRows
        
            For j = 2 To .MaxCols Step 2
                If FunValueCheck(Gf_Get_Cell_Text(SpreadHeader + 2, j), Gf_Get_Cell_Text(i, SpreadHeader + 1), GF_GET_CELL_VALUE(i, j)) = False Then
                    Call subCellBackColor(j, i)
                    iCnt = iCnt + 1
                End If
            Next j
        
        Next i
    
        If iCnt > 0 Then
            sMesg = "Out of production constraint"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    End With

End Sub

Private Function FunValueCheck(ByVal sTHK_GRP As String, ByVal sWID_GRP As String, ByVal dValue As Double) As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    If UBound(arrValue, 2) < 0 Then Exit Function
    
    For i = 0 To UBound(arrValue, 2)
    
        'If dValue = 310 Then
         '   Exit Function
        'End If
        
        If arrValue(0, i) = sTHK_GRP And arrValue(1, i) = sWID_GRP Then
        
            If arrValue(2, i) >= 0 And arrValue(3, i) <> 0 Then
                
                If arrValue(2, i) <= dValue And dValue <= arrValue(3, i) Then
                    FunValueCheck = True
                Else
                    FunValueCheck = False
                End If
                Exit Function
                            
            End If
        
        End If
    Next i
    
    FunValueCheck = True
    
End Function

Private Sub subCellBackColor(ByVal iCol As Integer, ByVal iRow As Integer)
    
    With ss1
        .Col = iCol
        .Row = iRow
        .BackColor = vbRed
    End With
    
End Sub

Private Function GF_GET_CELL_VALUE(ByVal iRow As Long, ByVal iCol As Long) As Variant
    
    With ss1
        .Row = iRow
        .Col = iCol
        GF_GET_CELL_VALUE = IIf(.Value = "", 0, .Value)
    End With
    
End Function

Private Function Gf_Get_Cell_Text(ByVal iRow As Long, ByVal iCol As Long) As Variant
    
    With ss1
        .Row = iRow
        .Col = iCol
        Gf_Get_Cell_Text = .Text
    End With
    
End Function
