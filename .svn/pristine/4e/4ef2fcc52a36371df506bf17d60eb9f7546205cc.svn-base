VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form VD_MODIF 
   BackColor       =   &H00E0E0E0&
   Caption         =   "VD操作记录输入项目"
   ClientHeight    =   1935
   ClientLeft      =   615
   ClientTop       =   4845
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   13560
   Begin VB.TextBox txt_m_heat_no 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Text            =   " "
      Top             =   480
      Width           =   999
   End
   Begin VB.TextBox txt_m_ldno 
      Height          =   315
      Left            =   3735
      TabIndex        =   16
      Text            =   " "
      Top             =   480
      Width           =   750
   End
   Begin VB.TextBox txt_m_stl_wgt 
      Height          =   315
      Left            =   4530
      TabIndex        =   15
      Text            =   " "
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txt_m_arrv_tm 
      Height          =   315
      Left            =   5550
      TabIndex        =   14
      Text            =   " "
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox txt_m_sta_tm 
      Height          =   315
      Left            =   6825
      TabIndex        =   13
      Text            =   " "
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox txt_m_end_ts 
      Height          =   315
      Left            =   8115
      TabIndex        =   12
      Text            =   " "
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox txt_m_dep_ts 
      Height          =   315
      Left            =   9405
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox txt_m_arrv_temp 
      Height          =   315
      Left            =   10695
      TabIndex        =   10
      Text            =   " "
      Top             =   480
      Width           =   870
   End
   Begin VB.TextBox txt_m_sta_temp 
      Height          =   315
      Left            =   11595
      TabIndex        =   9
      Text            =   " "
      Top             =   480
      Width           =   870
   End
   Begin VB.TextBox txt_m_end_temp 
      Height          =   315
      Left            =   12495
      TabIndex        =   8
      Text            =   " "
      Top             =   480
      Width           =   870
   End
   Begin VB.TextBox txt_m_dep_temp 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   " "
      Top             =   1350
      Width           =   855
   End
   Begin VB.ComboBox cbo_m_bb_status 
      Height          =   315
      ItemData        =   "VD_MODIFY.frx":0000
      Left            =   1035
      List            =   "VD_MODIFY.frx":000D
      TabIndex        =   6
      Text            =   " "
      Top             =   1350
      Width           =   930
   End
   Begin VB.TextBox txt_m_elect 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2010
      TabIndex        =   5
      Text            =   " "
      Top             =   1350
      Width           =   1095
   End
   Begin VB.TextBox txt_m_commt 
      Height          =   315
      Left            =   3150
      TabIndex        =   4
      Text            =   " "
      Top             =   1350
      Width           =   6015
   End
   Begin VB.ComboBox P_GROUP 
      Height          =   315
      ItemData        =   "VD_MODIFY.frx":0024
      Left            =   7800
      List            =   "VD_MODIFY.frx":0026
      TabIndex        =   3
      Text            =   " "
      Top             =   1890
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox P_SHIFT 
      Height          =   315
      ItemData        =   "VD_MODIFY.frx":0028
      Left            =   7800
      List            =   "VD_MODIFY.frx":002A
      TabIndex        =   2
      Text            =   " "
      Top             =   1530
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_m_re_cd 
      Height          =   315
      Left            =   2940
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   750
   End
   Begin VB.TextBox txt_m_stlgrd 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   1725
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   150
      Width           =   999
      _ExtentX        =   1773
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   120
      Top             =   990
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      Caption         =   "离开温度"
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   1035
      Top             =   990
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "搅拌情况"
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
      Left            =   2010
      Top             =   990
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "最低真空度"
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   3150
      Top             =   990
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   556
      Caption         =   "备注"
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
   Begin InDate.UDate P_DATE 
      Height          =   375
      Left            =   8520
      TabIndex        =   18
      Top             =   1530
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   End
   Begin Threed.SSCommand cmd_ok 
      Height          =   435
      Left            =   11220
      TabIndex        =   19
      Top             =   1230
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
   End
   Begin Threed.SSCommand cmd_exit 
      Height          =   435
      Left            =   12360
      TabIndex        =   20
      Top             =   1230
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3735
      Top             =   150
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "钢包号"
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   4530
      Top             =   150
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "钢水量"
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   5550
      Top             =   150
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "到达时间"
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
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   6825
      Top             =   150
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "开始时间"
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
   Begin InDate.ULabel ULabel21 
      Height          =   315
      Left            =   8115
      Top             =   150
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "结束时间"
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
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Left            =   9405
      Top             =   150
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "离开时间"
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   10695
      Top             =   150
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      Caption         =   "到达温度"
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
   Begin InDate.ULabel ULabel24 
      Height          =   315
      Left            =   11595
      Top             =   150
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      Caption         =   "开始温度"
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
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   12495
      Top             =   150
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      Caption         =   "结束温度"
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
   Begin InDate.ULabel ULabel26 
      Height          =   315
      Left            =   2940
      Top             =   150
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "再处理"
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
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   1170
      Top             =   150
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
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
   End
End
Attribute VB_Name = "VD_MODIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      LF RESLT MODIFICATION
'-- Program ID        AFE5013C
'-- Document No
'-- Designer          H.M.G
'-- Coder             H.M.G
'-- Date              2003.7.23
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
Public sDateTime As String          'Active Form Authority Setting
Public sQuery_Rt As String          'Active Form Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection




Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"              'form类型
       Call Gp_Ms_Collection(txt_m_heat_no, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_m_re_cd, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_m_stlgrd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(P_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(P_SHIFT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(P_GROUP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_m_ldno, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_m_stl_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_m_arrv_tm, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_m_sta_tm, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_m_end_ts, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_m_dep_ts, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_m_arrv_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_m_sta_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_m_end_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_m_dep_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_m_bb_status, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_m_elect, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_m_commt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
     Mc1.Add Item:="AFE5013C.P_SREFER", Key:="P-R"
     Mc1.Add Item:="AFE5013C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

End Sub
 
Private Sub cmd_exit_Click()
   Call Form_Exit
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Gp_FormCenter(Me)
    
    Call Form_Define
  
'    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Screen.MousePointer = vbDefault
    
    With AFE5010C.ss2
        .Row = .ActiveRow
        .Col = 1:     txt_m_heat_no.Text = .Text
        .Col = 2:     txt_m_re_cd.Text = .Text
        .Col = 3:     txt_m_stlgrd.Text = .Text
        .Col = 4:     txt_m_ldno.Text = .Text
        .Col = 5:     txt_m_stl_wgt.Text = .Text
        .Col = 7:     txt_m_arrv_tm.Text = .Text
        .Col = 8:     txt_m_sta_tm.Text = .Text
        .Col = 9:     txt_m_end_ts.Text = .Text
        .Col = 10:    txt_m_dep_ts.Text = .Text
        .Col = 11:    txt_m_arrv_temp.Text = .Text
        .Col = 12:    txt_m_sta_temp.Text = .Text
        .Col = 13:    txt_m_end_temp.Text = .Text
        .Col = 14:    txt_m_dep_temp.Text = .Text
        .Col = 15:    cbo_m_bb_status.Text = .Text
        .Col = 17:    txt_m_elect.Text = .Text
        .Col = 18:    txt_m_commt.Text = .Text
    End With
    
    If Mid(sAuthority, 3, 1) <> "1" Then
       cmd_ok.Enabled = False
    ElseIf Mid(sAuthority, 3, 1) = "1" Then
       cmd_ok.Enabled = True
    End If

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
'    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
End Sub

Private Sub Cmd_Ok_Click()
  If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
'     Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
     Unload Me
    Call AFE5010C.Form_Ref
  End If
  
  
End Sub

Private Sub Form_Activate()
'    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing

'    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
End Sub

