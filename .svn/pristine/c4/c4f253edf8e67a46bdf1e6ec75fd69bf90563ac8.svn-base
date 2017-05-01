VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form CAST_BAK 
   BackColor       =   &H00E0E0E0&
   Caption         =   "烘燥时间输入_CAST_BAK"
   ClientHeight    =   1575
   ClientLeft      =   4680
   ClientTop       =   5730
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5640
   Begin VB.TextBox txt_tnd_no 
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
      Left            =   1500
      MaxLength       =   8
      TabIndex        =   3
      Top             =   585
      Width           =   1050
   End
   Begin VB.TextBox txt_m_heat_no 
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
      Left            =   405
      MaxLength       =   8
      TabIndex        =   0
      Top             =   585
      Width           =   1050
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   405
      Top             =   225
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "炉号"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   2595
      Top             =   225
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "烘燥开始时间"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3960
      Top             =   225
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "烘燥结束时间"
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
   Begin CSTextLibCtl.sitxEdit txt_sta_date 
      Height          =   315
      Left            =   2595
      TabIndex        =   1
      Top             =   585
      Width           =   1320
      _Version        =   262145
      _ExtentX        =   2328
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __:__:__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "__ __:__"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   "__ __:__"
      Justification   =   1
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
   End
   Begin CSTextLibCtl.sitxEdit txt_end_date 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   585
      Width           =   1320
      _Version        =   262145
      _ExtentX        =   2328
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __:__:__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "__ __:__"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   "__ __:__"
      Justification   =   1
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   1500
      Top             =   225
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "中包号"
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
   Begin Threed.SSCommand cmd_OK 
      Height          =   435
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
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
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   435
      Left            =   2820
      TabIndex        =   5
      Top             =   1080
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
End
Attribute VB_Name = "CAST_BAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      LF RESLT MODIFICATION
'-- Program ID        AFB2010C
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
         Call Gp_Ms_Collection(txt_m_heat_no, "p", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_tnd_no, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_sta_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_end_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
     Mc1.Add Item:="AFH5010C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

End Sub
Private Sub Cmd_Cancel_Click()
   Call Form_Exit
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Gp_FormCenter(Me)
    
    Call Form_Define
  
'    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Screen.MousePointer = vbDefault
  
    With AFH5010C.SS1
        .Row = .ActiveRow
        .Col = 3
        txt_m_heat_no.Text = .Text
        .Col = 16
        txt_tnd_no.Text = .Text
        .Col = 19
        txt_sta_date.RawData = Mid(.Text, 1, 2) + Mid(.Text, 4, 2) + Mid(.Text, 7, 2)
        .Col = 20
        txt_end_date.RawData = Mid(.Text, 1, 2) + Mid(.Text, 4, 2) + Mid(.Text, 7, 2)
    End With
  
  If Mid(sAuthority, 3, 1) <> "1" Then
     cmd_ok.Enabled = False
  ElseIf Mid(sAuthority, 3, 1) = "1" Then
     cmd_ok.Enabled = True
  End If

End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
'    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
 '   Call Gp_Ms_ControlLock(Mc1("pControl"), False)
End Sub
Private Sub Cmd_Ok_Click()
'
'  If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
'     Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'  End If
    Dim adocTemp As New ADODB.Command
    adocTemp.ActiveConnection = M_CN1
    
    adocTemp.CommandText = "UPDATE FP_CASTEQUIP SET STA_DATE = '" + txt_sta_date.RawData + "', END_DATE = '" + txt_end_date.RawData + "' Where HEAT_NO = '" + txt_m_heat_no + "' AND EQUIP_KIND = 'TND'"
    adocTemp.Execute
    
    adocTemp.CommandText = "Commit"
    adocTemp.Execute
    
  Call AFH5010C.Form_Ref
  Unload Me
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

'    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, AFH5010C.sAuthority)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Private Sub txt_end_date_DblClick()
   txt_end_date.RawData = Format(Now, "DDHHMM")
End Sub

Private Sub txt_sta_date_DblClick()
   txt_sta_date.RawData = Format(Now, "DDHHMM")
End Sub
