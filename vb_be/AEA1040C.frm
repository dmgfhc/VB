VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AEA1040C 
   Caption         =   "录入炼钢编制标准_AEA1040C"
   ClientHeight    =   9225
   ClientLeft      =   540
   ClientTop       =   2250
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleMode       =   0  'User
   ScaleWidth      =   15136.22
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter5 
      Height          =   9135
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   16113
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      PaneTree        =   "AEA1040C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter6 
         Height          =   8565
         Left            =   0
         TabIndex        =   6
         Top             =   570
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   15108
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AEA1040C.frx":0052
         Begin SSSplitter.SSSplitter SSSplitter8 
            Height          =   6165
            Left            =   7920
            TabIndex        =   9
            Top             =   2400
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   10874
            _Version        =   196609
            SplitterBarWidth=   2
            SplitterBarJoinStyle=   0
            SplitterBarAppearance=   0
            BorderStyle     =   0
            BackColor       =   14737632
            PaneTree        =   "AEA1040C.frx":00C4
            Begin Threed.SSPanel SSPanel3 
               Height          =   510
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Width           =   7245
               _ExtentX        =   12779
               _ExtentY        =   900
               _Version        =   196609
               BackColor       =   14737632
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin Threed.SSCheck Chk_ss3 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   13
                  Top             =   150
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   450
                  _Version        =   196609
                  Font3D          =   2
                  ForeColor       =   255
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "连浇炉次内炉次数编制标准"
               End
            End
            Begin FPSpread.vaSpread SS3 
               Height          =   5625
               Left            =   0
               TabIndex        =   15
               Top             =   540
               Width           =   7245
               _Version        =   393216
               _ExtentX        =   12779
               _ExtentY        =   9922
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
               MaxCols         =   10
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "AEA1040C.frx":0116
            End
         End
         Begin SSSplitter.SSSplitter SSSplitter7 
            Height          =   6165
            Left            =   0
            TabIndex        =   8
            Top             =   2400
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   10874
            _Version        =   196609
            SplitterBarWidth=   2
            SplitterBarJoinStyle=   0
            SplitterBarAppearance=   0
            BorderStyle     =   0
            BackColor       =   14737632
            PaneTree        =   "AEA1040C.frx":0864
            Begin Threed.SSPanel SSPanel2 
               Height          =   510
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   7860
               _ExtentX        =   13864
               _ExtentY        =   900
               _Version        =   196609
               BackColor       =   14737632
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin Threed.SSCheck Chk_ss2 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   11
                  Top             =   150
                  Width           =   1770
                  _ExtentX        =   3122
                  _ExtentY        =   450
                  _Version        =   196609
                  Font3D          =   2
                  ForeColor       =   255
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "炉次编制量标准"
               End
            End
            Begin FPSpread.vaSpread SS2 
               Height          =   5625
               Left            =   0
               TabIndex        =   14
               Top             =   540
               Width           =   7860
               _Version        =   393216
               _ExtentX        =   13864
               _ExtentY        =   9922
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
               MaxCols         =   12
               MaxRows         =   1
               Protect         =   0   'False
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "AEA1040C.frx":08B6
            End
         End
         Begin FPSpread.vaSpread SS1 
            Height          =   2340
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   15165
            _Version        =   393216
            _ExtentX        =   26749
            _ExtentY        =   4128
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
            MaxCols         =   15
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AEA1040C.frx":10BD
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   540
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   953
         _Version        =   196609
         BackColor       =   14737632
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox TXT_PLT_NAME 
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
            Left            =   1935
            TabIndex        =   4
            Tag             =   "工厂"
            Top             =   120
            Width           =   2925
         End
         Begin VB.TextBox txt_PRC_line 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   6405
            MaxLength       =   1
            TabIndex        =   3
            Tag             =   "机号"
            Top             =   120
            Width           =   420
         End
         Begin VB.TextBox txt_plt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   2
            Tag             =   "工厂"
            Top             =   120
            Width           =   465
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   180
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "工厂"
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   5130
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "机号"
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
         Begin Threed.SSCheck Chk_ss1 
            Height          =   285
            Left            =   13560
            TabIndex        =   5
            Top             =   150
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   2
            ForeColor       =   255
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "板坯尺寸标准"
         End
      End
   End
End
Attribute VB_Name = "AEA1040C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEA1040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          JIA NING
'-- Coder             JIA NING
'-- Date              2003.6.19
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

Dim SwMess As Boolean

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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection
Dim Sc3 As New Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_PLT, "p", "n", "m", " ", "r", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_PLT_NAME, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_prc_line, "p", "n", "m", " ", "r", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    '------------------------------------------------SS1-----------------------------------------------------------
    
     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AEA1040C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:="AEA1040C.P_REFER1", Key:="P-R"
    sc1.Add Item:="AEA1040C.P_ONEROW1", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="sc"
    '------------------------------------------------------SS2------------------------------------------------------
    
    Call Gp_Sp_Collection(ss2, 1, "p", "n", " ", "i", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, "p", "n", " ", "i", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, "P", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, "P", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", "i", "a", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AEA1040C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:="AEA1040C.P_REFER2", Key:="P-R"
    sc2.Add Item:="AEA1040C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    '------------------------------------------------------SS3------------------------------------------------------
    
    Call Gp_Sp_Collection(ss3, 1, "p", "n", " ", "i", "a", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, "p", "n", " ", "i", "a", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, "P", "N", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", "", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", "", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", "", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", "", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AEA1040C.P_MODIFY3", Key:="P-M"
    Sc3.Add Item:="AEA1040C.P_REFER3", Key:="P-R"
    Sc3.Add Item:="AEA1040C.P_ONEROW3", Key:="P-O"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    '---------------------------------------------------------------------------------------------------------------
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Call Gp_Sp_ColHidden(ss1, 14, True)
    Call Gp_Sp_ColHidden(ss2, 2, True)
    Call Gp_Sp_ColHidden(ss2, 11, True)
    Call Gp_Sp_ColHidden(ss3, 2, True)
    Call Gp_Sp_ColHidden(ss3, 9, True)
    
    ss2.Enabled = False
    ss3.Enabled = False
    
End Sub

Private Sub Chk_ss1_Click(Value As Integer)

    Dim TT

    If Chk_ss1.Value = ssCBUnchecked Then
       If Chk_ss2.Value = ssCBUnchecked And Chk_ss3.Value = ssCBUnchecked Then
            Chk_ss1.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Gf_Sp_Change(Proc_Sc, sc1) Then
   
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss3.ForeColor = &H808080
        Chk_ss2.Value = ssCBUnchecked
        Chk_ss3.Value = ssCBUnchecked
        ss2.Enabled = False
        ss3.Enabled = False
        ss1.Enabled = True
        
    Else
        Select Case Proc_Sc("Sc").Item("Spread").Name
           Case "SS2"
               Chk_ss1.Value = ssCBUnchecked
               Chk_ss2.Value = ssCBChecked
               Chk_ss3.Value = ssCBUnchecked
               ss1.Enabled = False
               ss3.Enabled = False
               ss2.Enabled = True
               
           Case "SS3"
               Chk_ss1.Value = ssCBUnchecked
               Chk_ss2.Value = ssCBUnchecked
               Chk_ss3.Value = ssCBChecked
               ss1.Enabled = False
               ss3.Enabled = True
               ss2.Enabled = False
        End Select
        
    End If
        

End Sub

Private Sub Chk_ss2_Click(Value As Integer)

   If Chk_ss2.Value = ssCBUnchecked Then
       If Chk_ss1.Value = ssCBUnchecked And Chk_ss3.Value = ssCBUnchecked Then
            Chk_ss2.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Gf_Sp_Change(Proc_Sc, sc2) Then
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.ForeColor = &H808080
        Chk_ss3.ForeColor = &H808080
        Chk_ss1.Value = ssCBUnchecked
        Chk_ss3.Value = ssCBUnchecked
        ss1.Enabled = False
        ss2.Enabled = True
        ss3.Enabled = False
        
    Else
         Select Case Proc_Sc("Sc").Item("Spread").Name
            Case "SS1"
                Chk_ss2.Value = ssCBUnchecked
                Chk_ss1.Value = ssCBChecked
                Chk_ss3.Value = ssCBUnchecked
                 ss1.Enabled = True
                 ss2.Enabled = False
                 ss3.Enabled = False
                 
                
            Case "SS3"
                Chk_ss1.Value = ssCBUnchecked
                Chk_ss2.Value = ssCBUnchecked
                Chk_ss3.Value = ssCBChecked
                 ss1.Enabled = False
                 ss2.Enabled = False
                 ss3.Enabled = True
        End Select
        
    End If
    
End Sub

Private Sub Chk_ss3_Click(Value As Integer)

    If Chk_ss3.Value = ssCBUnchecked Then
       If Chk_ss1.Value = ssCBUnchecked And Chk_ss2.Value = ssCBUnchecked Then
            Chk_ss3.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Gf_Sp_Change(Proc_Sc, Sc3) Then
        Chk_ss3.ForeColor = &HFF&
        Chk_ss1.ForeColor = &H808080
        Chk_ss2.ForeColor = &H808080
        Chk_ss1.Value = ssCBUnchecked
        Chk_ss2.Value = ssCBUnchecked
          ss1.Enabled = False
          ss2.Enabled = False
          ss3.Enabled = True
    Else
        Select Case Proc_Sc("Sc").Item("Spread").Name
            Case "SS1"
                Chk_ss3.Value = ssCBUnchecked
                Chk_ss1.Value = ssCBChecked
                Chk_ss2.Value = ssCBUnchecked
                  ss1.Enabled = True
                 ss2.Enabled = False
                 ss3.Enabled = False
                 
            Case "SS2"
                Chk_ss1.Value = ssCBUnchecked
                Chk_ss3.Value = ssCBUnchecked
                Chk_ss2.Value = ssCBChecked
                
                 ss1.Enabled = False
                 ss2.Enabled = True
                 ss3.Enabled = False
        End Select
        
    End If

    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
     Chk_ss1.Value = ssCBChecked
     Chk_ss2.Value = ssCBUnchecked
     Chk_ss3.Value = ssCBUnchecked
     Chk_ss1.ForeColor = &HFF&
     Chk_ss2.ForeColor = &H808080
     Chk_ss3.ForeColor = &H808080

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
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    Call Gp_Sp_Setting(Sc3.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    
    Call Gp_Spl_SizeGet(SSSplitter6, "E-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "E-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(sc2.Item("Spread"), 4)
    Call Gp_Sp_HdColColor(Sc3.Item("Spread"), 4)
    
    txt_PLT.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    txt_prc_line.Text = "1"

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter6, "E-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "E-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) And Gf_Sp_Cls(sc2) And Gf_Sp_Cls(Sc3) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
        txt_PLT.Text = "B1"
        Call txt_plt_KeyUp(0, 0)
        txt_prc_line.Text = "1"
    End If
   
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Or _
       Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False) Or _
       Gf_Sp_Refer(M_CN1, Sc3, Mc1, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Exit Sub
    End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
'------------------------------------------------------------------
    If Chk_ss1.Value = -1 Then
        If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
          'Call Gp_Ms_Cls(Mc1("rControl"))
          'Call Gf_Sp_Cls(Sc2)
          'Call Gf_Sp_Cls(Sc3)
          Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        End If
    
    End If
    
    If Chk_ss2.Value = -1 Then
       If Gf_Sp_Process(M_CN1, sc2, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    If Chk_ss3.Value = -1 Then
       If Gf_Sp_Process(M_CN1, Sc3, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

End Sub

Public Sub Form_Ins()

    If Chk_ss1.Value = ssCBUnchecked And Chk_ss2.Value = ssCBUnchecked And Chk_ss3.Value = ssCBUnchecked Then
       MsgBox "first,you must chick the chickbox"
       Exit Sub
    
    End If
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    
    If Chk_ss1.Value = ssCBChecked Then
       Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 14)
    End If
    
    If Chk_ss2.Value = ssCBChecked Then
       Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If
    
    If Chk_ss3.Value = ssCBChecked Then
       Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
    If Chk_ss1.Value = ssCBChecked Then
       Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 14)
    End If
    
    If Chk_ss2.Value = ssCBChecked Then
       Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If
    
    If Chk_ss3.Value = ssCBChecked Then
       Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
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

'------------------------------------------------------ss1-------------------------------------------------------------
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
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 14)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 14)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If Chk_ss1.Value <> -1 Then
       MsgBox "PLEASE CHICK THE CHICKBOX1 ,FIRST"
        Exit Sub
    End If

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyF4 Then
    
        Select Case ss1.ActiveCol
        
            Case 1
            
               Set DD.sPname = Me.ss1
                    
                DD.sWitch = "SP"
                DD.sKey = "C0001"
                DD.rControl.Add Item:=1
                DD.rControl.Add Item:=2
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
                
        End Select
        
    End If
    
End Sub

Private Sub SS1_Change(ByVal Col As Long, ByVal Row As Long)
'Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim dMin  As Double
    Dim dMax  As Double
    Dim Dcurr As Double
    Dim BIJIAO As Double
    Dim iRow As Integer
    
    If ss1.MaxRows < 1 Or Row < 0 Or Row = 0 Or Col = 0 Then Exit Sub
    
    With ss1
        
        .Row = Row
        .Col = 0
        
        Select Case Col
        
         Case 5, 7, 9, 11
        
            .Row = Row
            .Col = Col
            If .CellTag = "False" Then Exit Sub
            
            If .Value = "" Then
                Dcurr = 0
            Else
                Dcurr = .Value
            End If
            
            .Row = Row
            .Col = Col - 1
            
            If .Value = "" Then
                BIJIAO = 0
            Else
                BIJIAO = .Value
            End If
            
            If BIJIAO > Dcurr Then
                .Row = Row
                .Col = Col
                .CellTag = "False"
                
                Call Gp_MsgBoxDisplay(" the wrong that  Min > Max happened")
                
                .Row = Row
                .Col = Col
                .CellTag = ""
                
                .Text = 0
                .Action = SS_ACTION_ACTIVE_CELL
                .EditMode = True
                .SetFocus
                Exit Sub
            End If
         End Select
                
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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

'-------------------------------------------------------------ss2---------------------------------------------------------------------------------

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_KeyUp(KeyCode As Integer, Shift As Integer)

    If Chk_ss2.Value <> -1 Then
       MsgBox "PLEASE CHICK THE CHICKBOX2 ,FIRST"
    
        Exit Sub
    
    End If

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyF4 Then
    
        Select Case ss2.ActiveCol
        
            Case 1
            
                Set DD.sPname = Me.ss2
                    
                DD.sWitch = "SP"
                DD.sKey = "C0001"
                DD.rControl.Add Item:=1
                DD.rControl.Add Item:=2
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
                
           Case 4
            
                Set DD.sPname = Me.ss2
                     
                DD.sWitch = "SP"
                DD.sKey = "E0001"
                DD.rControl.Add Item:=4
                DD.rControl.Add Item:=5
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
                
        End Select
        
    End If
    
End Sub

Private Sub SS2_Change(ByVal Col As Long, ByVal Row As Long)
'Private Sub ss2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim dMin  As Double
    Dim dMax  As Double
    Dim Dcurr As Double
    Dim BIJIAO As Double
    Dim iRow As Integer
    
    If ss2.MaxRows < 1 Or Row < 0 Or Row = 0 Or Col = 0 Then Exit Sub
    
    With ss2
        .Row = Row
        .Col = 0
        
        Select Case Col
        
            Case 7
               .Row = Row
               .Col = Col
               If .CellTag = "False" Then Exit Sub
               
               If .Value = "" Then
                   Dcurr = 0
               Else
                   Dcurr = .Value
               End If
               
               .Row = Row
               .Col = Col - 1
               
               If .Value = "" Then
                   BIJIAO = 0
               Else
                   BIJIAO = .Value
               End If
               
               If BIJIAO > Dcurr Then
                   .Row = Row
                   .Col = Col
                   .CellTag = "False"
                   
                   Call Gp_MsgBoxDisplay(" the wrong that  Min > Max happened")
                   
                   .Row = Row
                   .Col = Col
                   .CellTag = ""
                   
                   .Value = 0
                   .Action = SS_ACTION_ACTIVE_CELL
                   .EditMode = True
                   .SetFocus
                   Exit Sub
               End If
               
         End Select
                
    End With
    
End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

'----------------------------------------------------------ss3-----------------------------------------------------------------------------------

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub

Private Sub ss3_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
       Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
       Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss3_KeyUp(KeyCode As Integer, Shift As Integer)

    If Chk_ss3.Value <> -1 Then
        MsgBox "PLEASE CHICK THE CHICKBOX3 ,FIRST"
        Exit Sub
    End If

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyF4 Then
    
        Select Case ss3.ActiveCol
    
              Case 1
              
                Set DD.sPname = Me.ss3
                     
                DD.sWitch = "SP"
                DD.sKey = "C0001"
                DD.rControl.Add Item:=1
                DD.rControl.Add Item:=2
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Case 4
            
                Set DD.sPname = Me.ss3
                DD.sWitch = "SP"
                DD.sKey = "Q0048"
                DD.rControl.Add Item:=4
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
                
        End Select
        
    End If
    
End Sub

Private Sub SS3_Change(ByVal Col As Long, ByVal Row As Long)
'Private Sub SS3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim dMin  As Double
    Dim dMax  As Double
    Dim Dcurr As Double
    Dim BIJIAO As Double
    Dim iRow As Integer
    
    With ss3
        
        If ss3.MaxRows < 1 Or Row < 0 Or Row = 0 Or Col = 0 Then Exit Sub
        .Row = Row
        .Col = 0
        
        Select Case Col
        
         Case 6
            .Row = Row
            .Col = Col
            
            If .CellTag = "False" Then Exit Sub
            
            If .Value = "" Then
                Dcurr = 0
            Else
                Dcurr = .Value
            End If
            
            .Row = Row
            .Col = Col - 1
            If .Value = "" Then
                BIJIAO = 0
            Else
                BIJIAO = .Value
            End If
            
            If BIJIAO > Dcurr Then
                .Row = Row
                .Col = Col
                .CellTag = "False"
                
                Call Gp_MsgBoxDisplay(" the wrong that  Min > Max happened")
                
                .Row = Row
                .Col = Col
                .CellTag = ""
                
                .Value = 0
                .Action = SS_ACTION_ACTIVE_CELL
                .EditMode = True
                
                .SetFocus
                Exit Sub
            End If
            
         End Select
                
    End With
    
End Sub

Private Sub ss3_LostFocus()
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss3_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss3
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_PLT
        DD.rControl.Add Item:=txt_PLT_NAME
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_PLT.Text)) = txt_PLT.MaxLength Then
        txt_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_PLT.Text), 2)
    Else
        txt_PLT_NAME.Text = ""
    End If

End Sub
