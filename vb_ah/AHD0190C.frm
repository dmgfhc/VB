VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AHD0190C 
   Caption         =   "����ʵ����ѯ_AHD0190C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_plate_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6555
      MaxLength       =   14
      TabIndex        =   19
      Tag             =   "���Ϻ�"
      Top             =   540
      Width           =   2055
   End
   Begin VB.TextBox txt_out_sheet_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      MaxLength       =   14
      TabIndex        =   18
      Tag             =   "������"
      Top             =   540
      Width           =   1785
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   7935
      MaxLength       =   2
      TabIndex        =   17
      Top             =   165
      Width           =   675
   End
   Begin VB.TextBox txt_shp_ist_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10455
      MaxLength       =   11
      TabIndex        =   16
      Tag             =   "����ָʾ��"
      Top             =   540
      Width           =   2145
   End
   Begin VB.TextBox txt_sale_way 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10455
      MaxLength       =   2
      TabIndex        =   15
      Tag             =   "����"
      Top             =   915
      Width           =   465
   End
   Begin VB.TextBox txt_sale_way_name 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10920
      MaxLength       =   40
      TabIndex        =   14
      Tag             =   "Dept_Cd_Name"
      Top             =   915
      Width           =   1680
   End
   Begin VB.TextBox txt_ord_no 
      Height          =   315
      Left            =   6555
      MaxLength       =   11
      TabIndex        =   10
      Tag             =   "������"
      Top             =   915
      Width           =   1485
   End
   Begin VB.TextBox txt_prod_cd 
      Height          =   315
      Left            =   6555
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "��Ʒ"
      Top             =   165
      Width           =   435
   End
   Begin VB.TextBox txt_cust_nm 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2670
      MaxLength       =   40
      TabIndex        =   4
      Top             =   915
      Width           =   2205
   End
   Begin VB.TextBox txt_cust_cd 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      MaxLength       =   6
      TabIndex        =   3
      Top             =   915
      Width           =   1050
   End
   Begin VB.TextBox txt_ord_item 
      Height          =   315
      Left            =   8055
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "�������к�"
      Top             =   915
      Width           =   555
   End
   Begin VB.TextBox txt_cur_inv 
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
      Left            =   10830
      TabIndex        =   1
      Top             =   165
      Width           =   1770
   End
   Begin VB.TextBox txt_cur_inv_code 
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
      Left            =   10455
      MaxLength       =   2
      TabIndex        =   0
      Top             =   165
      Width           =   375
   End
   Begin InDate.UDate dtp_out_to 
      Height          =   310
      Left            =   3435
      TabIndex        =   6
      Tag             =   "��������"
      Top             =   170
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
   End
   Begin InDate.ULabel ULabel8 
      Height          =   300
      Left            =   3090
      Top             =   180
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Caption         =   "��"
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
   End
   Begin InDate.UDate dtp_out_fr 
      Height          =   310
      Left            =   1605
      TabIndex        =   7
      Tag             =   "��������"
      Top             =   170
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
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   5160
      Top             =   915
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      Caption         =   "������/���к�"
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
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   195
      Top             =   165
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "��������"
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
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   195
      Left            =   15
      TabIndex        =   8
      Top             =   165
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      _Version        =   196609
      ForeColor       =   255
      Value           =   1
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   195
      Left            =   15
      TabIndex        =   9
      Top             =   525
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      _Version        =   196609
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   195
      Top             =   915
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "�ͻ�"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   9045
      Top             =   915
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "���۹�˾"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   5160
      Top             =   165
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      Caption         =   "��Ʒ"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   7440
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1650
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   13123
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
      MaxCols         =   27
      MaxRows         =   1
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AHD0190C.frx":0000
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   12465
      Top             =   1305
      Visible         =   0   'False
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
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   9795
      Top             =   1305
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "��Ʒ����"
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
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   9045
      Top             =   165
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "�ֿ�"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   9045
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "�ᵥ��"
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
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   7050
      Top             =   165
      Width           =   870
      _ExtentX        =   1535
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   195
      Top             =   540
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   5160
      Top             =   540
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      Caption         =   "���Ϻ�"
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
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   13815
      X2              =   15000
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   11145
      X2              =   12330
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Label lbl_prod_wgt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   13830
      TabIndex        =   13
      Top             =   1305
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lbl_prod_num 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   11145
      TabIndex        =   12
      Top             =   1305
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "AHD0190C"
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
'-- Program ID        AHD0020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.19
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
       Call Gp_Ms_Collection(dtp_out_fr, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_out_to, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_cd, "P", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_cust_cd, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_cust_nm, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ord_no, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ord_item, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_cur_inv, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_sale_way, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_sale_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_shp_ist_no, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_out_sheet_no, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_plate_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
 '   Sc1.Add Item:="AHD0020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AHD0190C.P_REFER", Key:="P-R"
 '   Sc1.Add Item:="AHD0020C.P_ONEROW", Key:="P-O"
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
    'Duplicate Count
'    iDupCnt = 1

    'Sum Column Count
    iSumCnt = 3

    'Sum Column Setting

    iSumCol.Add Item:=20
    iSumCol.Add Item:=21
    iSumCol.Add Item:=19
    
End Sub

Private Sub dtp_out_fr_GotFocus()

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Menu_Setting

'    txt_ord_no.Enabled = False
'    txt_ord_item.Enabled = False
'    dtp_out_fr.Enabled = True
'    dtp_out_to.Enabled = True
'    txt_prod_cd.Enabled = True
'    txt_cust_cd.Enabled = True
'    txt_cust_nm.Enabled = True
'    text_cur_inv_code.Enabled = True
'    text_cur_inv.Enabled = True
'    dtp_out_fr.RawData = ""
'    dtp_out_to.RawData = ""
    
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

'    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

    txt_prod_cd.Text = "PP"

'    Chk_ss1.Value = ssCBChecked
'    Chk_ss2.Value = ssCBUnchecked
     
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
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set iSumCol = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Menu_Setting
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
     '   rControl(1).SetFocus
    End If
    dtp_out_fr.RawData = ""
    dtp_out_to.RawData = ""
    txt_cust_nm.Text = ""
'    lbl_prod_wgt.Caption = ""
'    lbl_prod_num.Caption = ""
    txt_cur_inv.Text = ""
    txt_prod_cd.Text = "PP"
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sQuery As String
    Dim sMesg As String
    
    If Trim(txt_prod_cd.Text) = "" Then
        Call Gp_MsgBoxDisplay("��Ʒ��������...")
        Exit Sub
    End If
    
    sQuery = "{call AHD0190C.P_REFER ('" + dtp_out_fr.RawData + "','" + dtp_out_to.RawData + "','" + txt_prod_cd.Text + "','" + txt_cust_cd.Text + "','" + txt_ord_no.Text + "','" + txt_ord_item.Text + "','" + txt_cur_inv_code.Text + "','" + txt_sale_way.Text + "','" + txt_shp_ist_no.Text + "','" + txt_plt.Text + "','" + txt_out_sheet_no.Text + "','" + txt_plate_no.Text + "')}"
    
'    If Chk_ss2.Value = ssCBUnchecked Then
'       If dtp_out_fr.RawData <> "" And dtp_out_to.RawData <> "" Then
'           If Gf_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
'           If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
'           If Gf_Total_Display(M_CN1, Proc_Sc("SC"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
'              Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Menu_Setting
'           End If
'             ss1.Row = ss1.MaxRows
'
'             If ss1.MaxRows > 1 Then
'                 ss1.Col = 10
'                 lbl_prod_num.Caption = ss1.Text
'                 ss1.Col = 11
'                 lbl_prod_wgt.Caption = ss1.Text
'             Else
'                 lbl_prod_num.Caption = ""
'                 lbl_prod_wgt.Caption = ""
'
'             End If
'       Else
'          Call Gp_MsgBoxDisplay("�������ڱ�������")
'          Exit Sub
'       End If
'    End If
'
'
'    If Chk_ss1.Value = ssCBUnchecked Then
'        If txt_ORD_NO.Text <> "" Then
'           If Gf_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
           If Gf_Total_Display(M_CN1, Proc_Sc("SC"), sQuery, 0, iSumCnt, iSumCol) Then
              
              Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
              Call Menu_Setting
           End If
'             ss1.Row = ss1.MaxRows
'
'             If ss1.MaxRows > 1 Then
'                 ss1.Col = 10
'                 lbl_prod_num.Caption = ss1.Text
'                 ss1.Col = 11
'                 lbl_prod_wgt.Caption = ss1.Text
'             Else
'                 lbl_prod_num.Caption = ""
'                 lbl_prod_wgt.Caption = ""
'
'             End If

'        Else
'           Call Gp_MsgBoxDisplay("�����ű�������")
'        End If
'    End If
               
               
Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    
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
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

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
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

'Private Sub Chk_ss1_Click(Value As Integer)
'
'    If Chk_ss1.Value = ssCBUnchecked Then
'       If Chk_ss2.Value = ssCBUnchecked Then
'
'            Chk_ss1.Value = ssCBChecked
'            txt_ord_no.Enabled = False
'            txt_ord_item.Enabled = False
'            txt_ord_no.Text = ""
'            txt_ord_item.Text = ""
'            dtp_out_fr.Enabled = True
'            dtp_out_to.Enabled = True
'            txt_prod_cd.Enabled = True
'            txt_cust_cd.Enabled = True
'            txt_cust_nm.Enabled = True
'            text_cur_inv_code.Enabled = True
'            text_cur_inv.Enabled = True
'       End If
'       Exit Sub
'
'    End If
'
'    Chk_ss1.ForeColor = &HFF&
'    Chk_ss2.ForeColor = &H808080
'    Chk_ss2.Value = ssCBUnchecked
'    txt_ord_no.Enabled = False
'    txt_ord_item.Enabled = False
'    txt_ord_no.Text = ""
'    txt_ord_item.Text = ""
'    dtp_out_fr.Enabled = True
'    dtp_out_to.Enabled = True
'    txt_prod_cd.Enabled = True
'    txt_cust_cd.Enabled = True
'    txt_cust_nm.Enabled = True
'    text_cur_inv_code.Enabled = True
'    text_cur_inv.Enabled = True
'
'End Sub

'Private Sub Chk_ss2_Click(Value As Integer)
'
'    If Chk_ss2.Value = ssCBUnchecked Then
'       If Chk_ss1.Value = ssCBUnchecked Then
'            Chk_ss2.Value = ssCBChecked
'            txt_ord_no.Enabled = True
'            txt_ord_item.Enabled = True
'            dtp_out_fr.Enabled = False
'            dtp_out_to.Enabled = False
'            txt_prod_cd.Enabled = False
'            txt_cust_cd.Enabled = False
'            txt_cust_nm.Enabled = False
'            text_cur_inv_code.Enabled = False
'            text_cur_inv.Enabled = False
'            dtp_out_fr.RawData = ""
'            dtp_out_to.RawData = ""
'            txt_prod_cd.Text = ""
'            txt_cust_cd.Text = ""
'            txt_cust_nm.Text = ""
'
'        End If
'        Exit Sub
'
'    End If
'
'    Chk_ss1.ForeColor = &H808080
'    Chk_ss2.ForeColor = &HFF&
'    Chk_ss1.Value = ssCBUnchecked
'    txt_ord_no.Enabled = True
'    txt_ord_item.Enabled = True
'    dtp_out_fr.Enabled = False
'    dtp_out_to.Enabled = False
'    txt_prod_cd.Enabled = False
'    txt_cust_cd.Enabled = False
'    txt_cust_nm.Enabled = False
'    text_cur_inv_code.Enabled = False
'    text_cur_inv.Enabled = False
'    dtp_out_fr.RawData = ""
'    dtp_out_to.RawData = ""
'    txt_prod_cd.Text = ""
'    txt_cust_cd.Text = ""
'    txt_cust_nm.Text = ""
'
'
'End Sub


Public Sub Menu_Setting()

    MDIMain.MenuTool.Buttons(4).Enabled = False    'Save
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel
    
End Sub


Private Sub TXT_CUST_CD_DblClick()

     Call TXT_CUST_CD_KeyUp(vbKeyF4, 0)
     
End Sub

Private Sub TXT_CUST_CD_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
End Sub

Private Sub TXT_CUST_CD_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_nm

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    
    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_nm.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_nm.Text = ""
    End If

End Sub

Private Sub txt_ord_no_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
End Sub

Private Sub txt_prod_cd_DblClick()

     Call txt_prod_cd_KeyUp(vbKeyF4, 0)
     
End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    

End Sub

Private Sub txt_cur_inv_code_DblClick()

    Call txt_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cur_inv_code_Change()
    If Len(Trim(txt_cur_inv_code.Text)) = txt_cur_inv_code.MaxLength Then
        txt_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv_code.Text, 2)
    Else
        txt_cur_inv.Text = ""
    End If

End Sub

Private Sub txt_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=txt_cur_inv_code
        DD.rControl.Add Item:=txt_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
End Sub

Private Sub TXT_SALE_WAY_DblClick()

    Call txt_sale_way_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_SALE_WAY_Change()
    If Len(Trim(txt_sale_way.Text)) = txt_sale_way.MaxLength Then
        txt_sale_way_name.Text = Gf_ComnNameFind(M_CN1, "B0010", txt_sale_way.Text, 2)
    Else
        txt_sale_way_name.Text = ""
    End If

End Sub

Private Sub txt_sale_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0010"
        DD.rControl.Add Item:=txt_sale_way
        DD.rControl.Add Item:=txt_sale_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_sale_way)) = txt_sale_way.MaxLength Then
        txt_sale_way_name.Text = Gf_ComnNameFind(M_CN1, "B0010", Trim(txt_sale_way.Text), 2)
    Else
        txt_sale_way_name.Text = ""
    End If

End Sub

Private Sub TXT_PLT_Change()

    If txt_plt.Text = "C3" Then
       txt_prod_cd.Text = "PP"
    End If

End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub


