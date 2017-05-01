VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGF3020C 
   Caption         =   "卷筒使用实绩及库存查询界面_AGF3020C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "卷筒使用实绩"
      TabPicture(0)   =   "AGF3020C.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSSplitter1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "卷筒库存查询"
      TabPicture(1)   =   "AGF3020C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSSplitter2"
      Tab(1).ControlCount=   1
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   8505
         Left            =   120
         TabIndex        =   1
         Top             =   375
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   15002
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "AGF3020C.frx":0038
         Begin Threed.SSFrame SSFrame1 
            Height          =   555
            Left            =   15
            TabIndex        =   2
            Top             =   15
            Width           =   14700
            _ExtentX        =   25929
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.ComboBox CBO_ROLL_NO 
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
               Left            =   1320
               TabIndex        =   8
               Top             =   120
               Width           =   1365
            End
            Begin InDate.ULabel ULabel5 
               Height          =   315
               Left            =   2745
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "使用日期"
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
            Begin InDate.UDate SDT_TO_DATE 
               Height          =   315
               Left            =   5895
               TabIndex        =   3
               Tag             =   "终止日期"
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
            Begin InDate.UDate SDT_FROM_DATE 
               Height          =   315
               Left            =   4125
               TabIndex        =   4
               Tag             =   "起始日期"
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
            Begin InDate.ULabel ULabel16 
               Height          =   315
               Left            =   120
               Top             =   120
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               Caption         =   "卷筒号"
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
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "～"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   5640
               TabIndex        =   5
               Top             =   120
               Width           =   255
            End
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   7860
            Left            =   15
            TabIndex        =   10
            Top             =   630
            Width           =   14700
            _Version        =   393216
            _ExtentX        =   25929
            _ExtentY        =   13864
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
            MaxRows         =   499
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AGF3020C.frx":008A
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   8430
         Left            =   -74880
         TabIndex        =   6
         Top             =   435
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   14870
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "AGF3020C.frx":1D9F
         Begin Threed.SSFrame SSFrame2 
            Height          =   555
            Left            =   15
            TabIndex        =   7
            Top             =   15
            Width           =   14715
            _ExtentX        =   25956
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.ComboBox cbo_roll 
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
               ItemData        =   "AGF3020C.frx":1DF1
               Left            =   1440
               List            =   "AGF3020C.frx":1DF3
               TabIndex        =   11
               Tag             =   "炉座号"
               Top             =   120
               Width           =   1425
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   120
               Top             =   120
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               Caption         =   "卷筒号"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   7785
            Left            =   15
            TabIndex        =   9
            Top             =   630
            Width           =   14715
            _Version        =   393216
            _ExtentX        =   25956
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
            MaxCols         =   12
            MaxRows         =   51
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AGF3020C.frx":1DF5
         End
      End
   End
End
Attribute VB_Name = "AGF3020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      精整作业计划查询界面
'-- Program ID        AGG2060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2005.8.10
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
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting
Public QueryYN      As Boolean


Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

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

Dim SE, MPLATE_NO, plate_no As String

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lRow        As Long
Dim lRowRange   As Long
Private Sub Form_Define()


    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDT_FROM_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDT_TO_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           
       
             Call Gp_Ms_Collection(cbo_roll, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       

   
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "P", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "P", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
  
       'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGF3020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGF3020C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGF3020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
'
'    sc1.Item("Spread").Col = 0
'    sc1.Item("Spread").Row = 0
'    sc1.Item("Spread").Text = "◎"
'
'     Me.KeyPreview = True
'     Me.BackColor = &HE0E0E0
    
        'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "P", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
    


    'Spread_Collection
    
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGF3020C.P_REFER2", Key:="P-R"
 
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

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

    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)

   
    
    Screen.MousePointer = vbDefault

    sQuery_load = "SELECT ROLL_NO FROM gp_roll WHERE  ROLL_NO LIKE  'J%' ORDER BY ROLL_NO "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

'     If SSTab1.Tab = 1 Then

     sQuery_load = "SELECT ROLL_NO FROM gp_roll WHERE  ROLL_NO LIKE  'J%' ORDER BY ROLL_NO "
     Call Gf_ComboAdd(M_CN1, cbo_roll, sQuery_load)

'    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
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
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
End Sub


Public Sub Form_Cls()

    If SSTab1.Tab = 0 Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            SDT_FROM_DATE.SetFocus
            SDT_FROM_DATE = ""
            SDT_TO_DATE = ""

        End If
    Else
        Call Gf_Sp_Cls(sc2)
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

Dim iRow  As Integer
Dim I, j, Scr_wgt, Hm_wgt, Steel_wgt As Integer
Dim Iron_Rec, Iront_Rec, Ir_Rec, I_c, Iron_Use, Back_Wgt, I_Rec As Double


On Error GoTo Refer_Err
Dim sCid As String


QueryYN = False

 If SSTab1.Tab = 0 Then

If SDT_FROM_DATE.RawData = "" And Trim(CBO_ROLL_NO.Text) = "" Then
       SDT_FROM_DATE.Text = Format(Now, "YYYY-MM") + "-01"
    End If
    If SDT_TO_DATE.RawData = "" And Trim(CBO_ROLL_NO.Text) = "" Then
       SDT_TO_DATE.Text = Format(Now, "YYYY-MM-DD")
    End If
   
    
 If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
   If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
         
          For I = 1 To ss1.MaxRows
            ss1.Col = 2
            ss1.Row = I
            sCid = ss1.Text
            If sCid <> "" Then
               ss1.Col = 1
               ss1.Text = sCid
            End If
        Next I
         
         
          With ss1
      
             For I = 1 To .MaxRows
                .Row = I
                .Col = 5
                Iron_Rec = Iron_Rec + Val(.Text)
               
               .Row = I
                .Col = 6
                I_Rec = I_Rec + Val(.Text)
             Next I


            .MaxRows = .MaxRows + 1
             .Row = .MaxRows
             For I = 1 To .MaxCols
                 .Col = I
                 .BackColor = "&HE6E6FF"
             Next I

             .Col = 1
             .Text = "合计"
             .Lock = True
             .Col = 4
             .Lock = True
             .Col = 5
             .Text = Str(Iron_Rec)
             .Lock = True
             .Col = 6
             .Text = Str(I_Rec)
             .Lock = True
             
        End With
    
    End If
 Else
        
        If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
            ss2.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            
           With ss2
      
             For I = 1 To .MaxRows
                
                .Row = I
                .Col = 3
                 Iront_Rec = Iront_Rec + Val(.Text)
                
                
                .Row = I
                .Col = 4
                 Ir_Rec = Ir_Rec + Val(.Text)
                
                .Row = I
                .Col = 5
                 Iron_Rec = Iron_Rec + Val(.Text)
               
                .Row = I
                .Col = 6
                 I_Rec = I_Rec + Val(.Text)
                
                .Row = I
                .Col = 9
                 I_c = I_c + Val(.Text)
                
             Next I


            .MaxRows = .MaxRows + 1
             .Row = .MaxRows
             For I = 1 To .MaxCols
                 .Col = I
                 .BackColor = "&HE6E6FF"
             Next I

               .Col = 1
               .Text = "合计"
               .Lock = True
               .Col = 3
               .Text = Str(Iront_Rec)
               .Lock = True
            
               .Col = 4
               .Text = Str(Ir_Rec)
               .Lock = True
               .Col = 5
               .Text = Str(Iron_Rec)
               .Lock = True
               .Col = 6
               .Text = Str(I_Rec)
               .Lock = True
             
               .Col = 9
               .Text = Str(I_c)
               .Lock = True
             
            End With
     
            
        
        End If
    End If
    
    Exit Sub


Refer_Err:

End Sub




Public Sub Form_Pro()
   Dim icount As Integer
   
    For icount = 1 To ss1.MaxRows

        Select Case Trim(Gf_Sp_RcvData(ss1, 0, icount))

           Case "Input", "Update"

             With ss1
             .Col = 4
             If Not Gp_DateCheck(.Text, "S") Then
                         Call Gp_MsgBoxDisplay("请输入正确的使用时间")
                         Exit Sub
                      End If
                  End With

        End Select

    Next icount

             
      If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
     ss1.OperationMode = OperationModeNormal
    Call Form_Ref
  End If
End Sub


Public Sub Form_Ins()

   If SSTab1.Tab = 0 Then
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.Col = 13
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID
    
     Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 1
    ss1.BackColor = &HC0FFFF
    
       Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT ROLL_NO  FROM GP_ROLL WHERE ROLL_NO LIKE 'J%' ORDER BY ROLL_NO ")
End If
End Sub


Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
    
End Sub
Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    ss1.Col = 13
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID
    
    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 1
    ss1.BackColor = &HC0FFFF
    
       Call Pf_ComboAdd(M_CN1, ss1, 1, "SELECT ROLL_NO  FROM GP_ROLL WHERE ROLL_NO LIKE 'J%' ORDER BY ROLL_NO ")
    ss1.OperationMode = OperationModeNormal

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

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub


Public Sub Spread_Del()

    If SSTab1.Tab = 0 Then
    Call Gp_Sp_Del(Proc_Sc("SC"))
   End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim sCid As String
If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        ss1.Col = 0
        ss1.Row = ss1.ActiveRow
        If ss1.Text = "Update" Then
            ss1.Col = 13
            ss1.Text = sUserID
        End If
        
        If Col = 1 Then
           ss1.Col = 1
           sCid = ss1.Text
           ss1.Col = 2
           ss1.Text = sCid
        End If
    End If
    
End Sub
Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub
Private Sub SDT_FROM_DATE_DblClick()
    If SDT_FROM_DATE.RawData = "" Then
     SDT_FROM_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_TO_DATE.RawData = "" Then
        SDT_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub
Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

 If Row <> 0 Then
    If Col = 4 Then
        ss1.Col = Col
        ss1.Row = Row
        If ss1.Lock = False Then
           ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
           ss1.Col = 0
           If ss1.Text <> "Input" And ss1.Text <> "Delete" Then
              ss1.Text = "Update"
           End If
        End If
        
        
        End If

    End If

End Sub



Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub

    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub
Public Function Pf_ComboAdd(Conn As ADODB.Connection, ss As vaSpread, Col As Integer, sQuery As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim AdoRs As ADODB.Recordset
    Dim sList As String
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Pf_ComboAdd = False: Exit Function
    End If
    
'    If ClsChk Then
'        Cbo.Clear
'    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If AdoRs.Fields(0) <> vbNull Then
                sList = sList & AdoRs.Fields(0) & vbTab
                'Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Pf_ComboAdd = True
    Else
        Pf_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    ss.Col = Col
    ss.TypeComboBoxList = sList
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Pf_ComboAdd = False

End Function
Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

  If ss1.ActiveCol = 1 Then
       ss1.Row = ss1.ActiveRow
       ss1.Col = ss1.ActiveCol
       If Len(Trim(ss1.Text)) = 7 Then
          Dim sQuery As String
          sQuery = "SELECT ROLL_WGT FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 7
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
          ss1.Col = 1
          sQuery = "SELECT ISSUETALLYNO FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 8
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          ss1.Col = 1
          sQuery = "SELECT MTRLNO FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 9
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          ss1.Col = 1
          sQuery = "SELECT ROLL_PRICE FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 10
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
       Else
          ss1.Col = 7
          ss1.Text = ""
          
         
          ss1.Col = 8
          ss1.Text = ""
          
      
          ss1.Col = 9
          ss1.Text = ""
          
      
          ss1.Col = 10
          ss1.Text = ""
          
    End If
  End If
End Sub
Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
 If ss1.ActiveCol = 1 Then
       ss1.Row = ss1.ActiveRow
       ss1.Col = ss1.ActiveCol
       If Len(Trim(ss1.Text)) = 7 Then
          Dim sQuery As String
          sQuery = "SELECT ROLL_WGT FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 7
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
          ss1.Col = 1
          sQuery = "SELECT ISSUETALLYNO FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 8
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          ss1.Col = 1
          sQuery = "SELECT MTRLNO FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 9
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)
          
          ss1.Col = 1
          sQuery = "SELECT ROLL_PRICE FROM GP_ROLL   WHERE ROLL_NO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 10
          ss1.Text = Val(Gf_FloatFind(M_CN1, sQuery))
          
       Else
          ss1.Col = 7
          ss1.Text = ""
          
         
          ss1.Col = 8
          ss1.Text = ""
          
      
          ss1.Col = 9
          ss1.Text = ""
          
      
          ss1.Col = 10
          ss1.Text = ""
          
    End If
  End If
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

