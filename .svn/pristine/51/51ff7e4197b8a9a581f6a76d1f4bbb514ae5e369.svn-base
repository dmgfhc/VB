VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form PopupMas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PopupMaster"
   ClientHeight    =   4995
   ClientLeft      =   5925
   ClientTop       =   4245
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_biz_area 
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
      Left            =   1710
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "TABLE_ID"
      Top             =   1800
      Width           =   465
   End
   Begin VB.TextBox txt_biz_area_name 
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
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "TABLE_ID"
      Top             =   1800
      Width           =   4830
   End
   Begin VB.TextBox txt_ins_time 
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
      Left            =   1725
      MaxLength       =   6
      TabIndex        =   7
      Top             =   4005
      Width           =   1185
   End
   Begin VB.TextBox txt_cd_mana_name 
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
      Left            =   1710
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "CD_MANA_NAME"
      Top             =   1395
      Width           =   6765
   End
   Begin VB.TextBox txt_cd_mana_no 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1710
      MaxLength       =   5
      TabIndex        =   0
      Tag             =   "CD_MANA_NO"
      Top             =   990
      Width           =   870
   End
   Begin VB.TextBox txt_cd_desc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1710
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2610
      Width           =   4245
   End
   Begin VB.TextBox txt_ins_emp 
      Enabled         =   0   'False
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
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   8
      Top             =   4410
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9120
      TabIndex        =   9
      Top             =   0
      Width           =   9180
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   825
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   1455
         BandCount       =   1
         _CBWidth        =   11250
         _CBHeight       =   825
         _Version        =   "6.7.8988"
         Child1          =   "MenuTool"
         MinHeight1      =   765
         Width1          =   2400
         NewRow1         =   0   'False
         Begin MSComctlLib.Toolbar MenuTool 
            Height          =   765
            Left            =   30
            TabIndex        =   11
            Top             =   30
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   1349
            ButtonWidth     =   1244
            ButtonHeight    =   1349
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList2"
            HotImageList    =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Clear"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line1"
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Save"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Delete"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line2"
                  Style           =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Copy"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Paste"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line3"
                  Style           =   4
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Exit"
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6255
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   45
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":0159
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":02E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":048A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":05CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":0725
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6300
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   45
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":08C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":0A55
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":0C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":0E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":0F53
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":10A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMaster.frx":1226
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_cd_len 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Tag             =   "CD_LEN"
      Top             =   2205
      Width           =   600
      _Version        =   262145
      _ExtentX        =   1058
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
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
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
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
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      Undo            =   0
      Data            =   0
   End
   Begin InDate.UDate dtp_ins_date 
      Height          =   315
      Left            =   1710
      TabIndex        =   6
      Top             =   3600
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   270
      Top             =   990
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "CD_MANA_NO"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   270
      Top             =   1395
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "CD_MANA_NAME"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   270
      Top             =   2205
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "CD_LEN"
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   270
      Top             =   2610
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "CD_DESC"
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
      Left            =   270
      Top             =   3600
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "INS_DATE"
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
      Left            =   270
      Top             =   4005
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "INS_TIME"
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
      Left            =   270
      Top             =   4410
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "INS_EMP"
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   270
      Top             =   1800
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "BIZ_AREA"
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
Attribute VB_Name = "PopupMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Template System
'-- Sub_System Name   Common
'-- Program Name      Popup Sheet Template
'-- Program ID        POPUPMASTER (Master)
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

Dim Mc1 As New Collection           'Master Collection

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "PopMaster"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary )", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_cd_mana_no, "p", "n", "m", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_cd_mana_name, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_biz_area, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    'Call Gp_Ms_Collection(txt_biz_area_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_cd_len, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_cd_desc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(dtp_ins_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ins_time, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
     Mc1.Add Item:="PKG_MASTER.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="PKG_MASTER.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
          
     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Activate()

    If Mc1("pControl").Item(1).Text = "" Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority("PopupShe")
    
    Call Popup_Menu_Setting
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_FormCenter(Me)
    
    Screen.MousePointer = vbDefault
    
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

    Call PopupShe.Form_Ref

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    MenuTool.Buttons(4).Enabled = False    'Delete
    MenuTool.Buttons(6).Enabled = False    'Copy
    MenuTool.Buttons(7).Enabled = False    'Paste
    
    pControl(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then MenuTool.Buttons(4).Enabled = False   'Delete
    
End Sub

Public Sub Form_Pro()
   
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        txt_ins_emp.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call Popup_Menu_Setting
        End If
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then
        Call Popup_Menu_Setting
    End If
    
End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "Clear"              'Clear
            Call Form_Cls
        Case "Save"               'Process
            Call Form_Pro
        Case "Delete"             'Delete
            Call Form_Del
        Case "Copy"               'Copy
            Call Master_Cpy
        Case "Paste"              'Paste
            Call Master_Pst
        Case "Exit"               'Exit
            Call Form_Exit
    End Select

End Sub

Private Sub txt_biz_area_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Z0001"
        DD.rControl.Add Item:=txt_biz_area
        DD.rControl.Add Item:=txt_biz_area_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_biz_area)) = txt_biz_area.MaxLength Then
        txt_biz_area_name.Text = Gf_ComnNameFind(M_CN1, "Z0001", Trim(txt_biz_area.Text), 2)
    Else
        txt_biz_area_name.Text = ""
    End If

End Sub

Public Sub Popup_Menu_Setting()

    Select Case Mid(sAuthority, 2, 3)
    
        Case "000"      'No Authority
            MenuTool.Buttons(3).Enabled = False                     'Save
            MenuTool.Buttons(4).Enabled = False                     'Delete
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste
        
        Case "001"      'Delete Authority
            MenuTool.Buttons(3).Enabled = False                     'Save
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste
        
        Case "010"      'Update Authority
            MenuTool.Buttons(4).Enabled = False                     'Delete
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste
        
        Case "011"      'Update, Delete Authority
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste
        
        Case "100"      'Insert Authority
            MenuTool.Buttons(4).Enabled = False                     'Delete
        
        Case "101"      'Insert, Delete Authority
        
        Case "110"      'Insert, Update Authority
            MenuTool.Buttons(4).Enabled = False                     'Delete
        
        Case "111"      'Insert, Update, Delete Authority
    
    End Select
    
End Sub

