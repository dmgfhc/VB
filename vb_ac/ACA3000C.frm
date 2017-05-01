VERSION 5.00
Begin VB.Form ACA3000C 
   Caption         =   "订单进程现状查询"
   ClientHeight    =   9090
   ClientLeft      =   360
   ClientTop       =   1605
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text_BB_ORD_NO 
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
      Left            =   990
      MaxLength       =   11
      TabIndex        =   17
      Top             =   525
      Width           =   1290
   End
   Begin VB.TextBox Text_BB_DOME_FL_mate 
      Height          =   300
      Left            =   13305
      TabIndex        =   16
      Top             =   390
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   310
      Left            =   4260
      Max             =   1
      Min             =   99
      TabIndex        =   15
      Top             =   527
      Value           =   1
      Width           =   285
   End
   Begin VB.TextBox Text_ORD_ITEM 
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
      Left            =   3735
      MaxLength       =   2
      TabIndex        =   14
      Top             =   525
      Width           =   540
   End
   Begin VB.TextBox text_vbCHECK 
      Height          =   345
      Left            =   12285
      TabIndex        =   6
      Top             =   75
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CheckBox Check_CP_ORD_REM_WGT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "过量生产"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10650
      TabIndex        =   9
      Top             =   547
      Width           =   1170
   End
   Begin VB.TextBox Text_BB_PROD_CD_mate 
      Height          =   315
      Left            =   14175
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox Text_BB_REC_STS_Name 
      Height          =   315
      Left            =   14130
      TabIndex        =   11
      Top             =   150
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox Text_BB_DEST_CD 
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
      Left            =   6195
      MaxLength       =   6
      TabIndex        =   5
      Top             =   525
      Width           =   900
   End
   Begin VB.TextBox Text_BB_REC_STS 
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
      Left            =   6195
      MaxLength       =   1
      TabIndex        =   2
      Top             =   90
      Width           =   900
   End
   Begin VB.TextBox Text_BB_DOME_FL 
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
      Left            =   3735
      MaxLength       =   1
      TabIndex        =   1
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox Text_BB_PROD_CD 
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
      Left            =   990
      MaxLength       =   2
      TabIndex        =   0
      Top             =   90
      Width           =   1290
   End
   Begin VB.PictureBox ULabel1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   5805
      ScaleHeight     =   30
      ScaleWidth      =   30
      TabIndex        =   18
      Top             =   3630
      Width           =   30
   End
   Begin VB.CheckBox Check_CP_DEL_DELAY 
      BackColor       =   &H00E0E0E0&
      Caption         =   "交货期延迟"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9135
      TabIndex        =   8
      Top             =   547
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.CheckBox Check_ord_END 
      BackColor       =   &H00E0E0E0&
      Caption         =   "完成对象"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   270
      Left            =   7710
      TabIndex        =   7
      Top             =   547
      Width           =   1170
   End
   Begin VB.PictureBox ss1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7830
      Left            =   90
      ScaleHeight     =   7770
      ScaleWidth      =   15090
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1110
      Width           =   15150
   End
   Begin VB.PictureBox ULabel9 
      BackColor       =   &H00E1E4CD&
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
      Left            =   90
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   90
      Width           =   855
   End
   Begin VB.PictureBox ULabel2 
      BackColor       =   &H00E1E4CD&
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
      Left            =   2790
      ScaleHeight     =   255
      ScaleWidth      =   840
      TabIndex        =   20
      Top             =   90
      Width           =   900
   End
   Begin VB.PictureBox ULabel3 
      BackColor       =   &H00E1E4CD&
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
      Left            =   7710
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   21
      Top             =   90
      Width           =   855
   End
   Begin VB.PictureBox Udate_BB_DEL_FR 
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
      Height          =   315
      Left            =   8595
      ScaleHeight     =   255
      ScaleWidth      =   1380
      TabIndex        =   3
      Tag             =   "INS_DATE"
      Top             =   90
      Width           =   1440
   End
   Begin VB.PictureBox UDate_BB_DEL_TO 
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
      Height          =   315
      Left            =   10320
      ScaleHeight     =   255
      ScaleWidth      =   1395
      TabIndex        =   4
      Tag             =   "INS_DATE"
      Top             =   90
      Width           =   1455
   End
   Begin VB.PictureBox ULabel4 
      BackColor       =   &H00E1E4CD&
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
      Left            =   5295
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   22
      Top             =   525
      Width           =   855
   End
   Begin VB.PictureBox ULabel6 
      BackColor       =   &H00E1E4CD&
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
      Left            =   2790
      ScaleHeight     =   255
      ScaleWidth      =   840
      TabIndex        =   23
      Top             =   525
      Width           =   900
   End
   Begin VB.PictureBox ULabel7 
      BackColor       =   &H00E1E4CD&
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
      Left            =   5295
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   24
      Top             =   90
      Width           =   855
   End
   Begin VB.PictureBox ULabel5 
      BackColor       =   &H00E1E4CD&
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
      Left            =   90
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   25
      Top             =   525
      Width           =   855
   End
   Begin VB.TextBox Text_BB_DEST_CD_mate 
      Height          =   315
      Left            =   14190
      TabIndex        =   13
      Top             =   990
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   75
      X2              =   15245
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   75
      X2              =   15245
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line1 
      X1              =   10065
      X2              =   10245
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "ACA3000C"
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
'-- Program Name      Master Sheet Template
'-- Program ID        MSHEET
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sCheck As String
Dim iCount As Integer





Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(Text_BB_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_BB_DOME_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_BB_REC_STS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Udate_BB_DEL_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(UDate_BB_DEL_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Text_BB_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_BB_DEST_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(text_vbCHECK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
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
     Call Gp_Sp_Collection(SS1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
   'Spread_Collection
    Sc1.Add Item:=SS1, Key:="Spread"

   'Sc1.Add Item:="PKG_MSHEET.P_MODIFY", Key:="P-M"

    Sc1.Add Item:="ACA1020C.P_SREFER", Key:="P-R"

   'Sc1.Add Item:="PKG_MSHEET.P_ONEROW", Key:="P-O"

    
       
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=SS1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub



Private Sub Check_ord_END_Click()
  If Check_ord_END.Value = 1 Then
    Text_BB_REC_STS.Text = 3

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
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
           



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
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
    Set Sc1 = Nothing
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
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If

        Udate_BB_DEL_FR.Text = ""
        UDate_BB_DEL_TO.Text = ""

        Check_ord_END.Value = 0
        Check_CP_DEL_DELAY.Value = 0
        Check_CP_ORD_REM_WGT.Value = 0
        Text_BB_PROD_CD_mate.Text = ""
        Text_BB_REC_STS_Name.Text = ""
        Text_BB_DEST_CD_mate.Text = ""
        Text_BB_DOME_FL_mate.Text = ""
        
        text_vbCHECK.Text = ""
        iCount = 0
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim smesg As String
    
    Dim S As String
        If Text_ORD_ITEM.Text <> "" Then
            If Len(Text_ORD_ITEM.Text) = 1 Then
    
             S = Text_ORD_ITEM.Text
             Text_ORD_ITEM.Text = "0" + S
             
            End If
        End If
    
    If Text_BB_ORD_NO.Text = "" Then
        Text_ORD_ITEM.Text = ""
    End If
    
    
     If Check_ord_END.Value = 1 And Check_CP_DEL_DELAY.Value = 0 And Check_CP_ORD_REM_WGT = 0 Then
            sCheck = " AND A.REC_STS = 3 "                                                       '"100"
        ElseIf Check_ord_END.Value = 0 And Check_CP_DEL_DELAY.Value = 1 And Check_CP_ORD_REM_WGT = 0 Then
            sCheck = " AND B.DEL_DELAY_DAY > 0 "                                                 '"010"
        ElseIf Check_ord_END.Value = 0 And Check_CP_DEL_DELAY.Value = 0 And Check_CP_ORD_REM_WGT = 1 Then
            sCheck = " AND B.ORD_REM_WGT < 0 "                                                   '"001 "
        ElseIf Check_ord_END.Value = 1 And Check_CP_DEL_DELAY.Value = 1 And Check_CP_ORD_REM_WGT = 0 Then
            sCheck = " AND A.REC_STS = 3 AND B.DEL_DELAY_DAY > 0 "                               '"110"
        ElseIf Check_ord_END.Value = 0 And Check_CP_DEL_DELAY.Value = 1 And Check_CP_ORD_REM_WGT = 1 Then
            sCheck = " AND B.DEL_DELAY_DAY > 0  AND B.ORD_REM_WGT < 0 "                          '"011"
        ElseIf Check_ord_END.Value = 1 And Check_CP_DEL_DELAY.Value = 0 And Check_CP_ORD_REM_WGT = 1 Then
            sCheck = " AND A.REC_STS = 3 AND B.ORD_REM_WGT < 0 "                                 '"101"
        ElseIf Check_ord_END.Value = 1 And Check_CP_DEL_DELAY.Value = 1 And Check_CP_ORD_REM_WGT = 1 Then
            sCheck = " AND A.REC_STS = 3 AND B.DEL_DELAY_DAY > 0  AND B.ORD_REM_WGT < 0 "        '"111"
        Else
            sCheck = ""                                                                          '"000"
        End If
            text_vbCHECK.Text = sCheck
    
    
    If UDate_BB_DEL_TO.RawData >= Udate_BB_DEL_FR.RawData Or UDate_BB_DEL_TO.RawData = "" Then
    
            If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
            smesg = Gf_Ms_NeceCheck(nControl)
              If smesg = "OK" Then
            
                    smesg = Gf_Ms_NeceCheck2(mControl)
                    If smesg = "OK" Then
                    
                        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
                            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                            Exit Sub
                        End If
                        
                    Else
                        smesg = smesg + " Must input according to length of item"
                        Call Gp_MsgBoxDisplay(smesg)
                    End If
               
               Else
                   smesg = smesg + " Must input necessarily"
                   Call Gp_MsgBoxDisplay(smesg)
                   
               End If
               
               Exit Sub
            
Refer_Err:
 Else
    Call MsgBox("输入日期不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
 End If
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    
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
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc")("Spread"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
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
        Set Active_Spread = Me.SS1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub




Private Sub Text_BB_DOME_FL_Change()
               Select Case Text_BB_DOME_FL.Text
                     Case "E", "D", ""
                Case Else
                      Text_BB_DOME_FL.Text = ""
                      
                      Call MsgBox("订单分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
                End Select
End Sub

Private Sub Text_BB_DOME_FL_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF4 Then
 
        DD.Switch = "MS"

        DD.sKey = "B0002"


        DD.rControl.Add Item:=Text_BB_DOME_FL
        DD.rControl.Add Item:=Text_BB_DOME_FL_mate
   
        DD.nameType = "2"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
       
        
        Call Gf_Common_DD(M_CN1, KeyCode)
       

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_DOME_FL.Text)) = Text_BB_DOME_FL.MaxLength Then
       '  Gf_ComnNAME_Find( 连接字符串, DD.sKEy内容 ,DD.nameType)
       ' Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        Text_BB_DOME_FL_mate.Text = Gf_ComnNameFind(M_CN1, "B0002", Text_BB_DOME_FL.Text, 2)
    Else
        Text_BB_DOME_FL_mate.Text = ""
    End If
End Sub

Private Sub Text_BB_ORD_NO_Change()

If Len(Text_BB_ORD_NO.Text) = Text_BB_ORD_NO.MaxLength Then

  
        Dim squery As String

        Dim AdoRs As adodb.Recordset


        squery = " SELECT NUM_ITEM FROM BP_ORDER WHERE ORD_NO = '" & Trim(Text_BB_ORD_NO.Text) & "'"
           
           Set AdoRs = New adodb.Recordset

           AdoRs.Open squery, M_CN1, adOpenKeyset
           If AdoRs.EOF Or AdoRs.BOF Then

              AdoRs.Close
              Set AdoRs = Nothing
                 Text_ORD_ITEM.Text = ""
                 iCount = 0
                 Call Gp_MsgBoxDisplay("无此订单！ ")
                 
              Exit Sub
           End If


           iCount = AdoRs.Fields(0)
           Text_ORD_ITEM.Text = "01"
           Exit Sub


End If
End Sub

Private Sub Text_BB_ORD_NO_LostFocus()
If Text_BB_ORD_NO.Text <> "" Then
   If (Len(Text_BB_ORD_NO.Text) < Text_BB_ORD_NO.MaxLength) Then
      Text_ORD_ITEM.Text = ""
      iCount = 0
      Call Gp_MsgBoxDisplay("订单号不符合规范！")
      Text_BB_ORD_NO.SetFocus
   End If
End If

End Sub

Private Sub Text_BB_PROD_CD_Change()
 If Len(Text_BB_PROD_CD.Text) = Text_BB_PROD_CD.MaxLength Then
                Select Case Text_BB_PROD_CD.Text
                     Case "SL", "PP", "HC", "", "**"
                Case Else
                      Text_BB_PROD_CD.Text = ""
                      
                      Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
                End Select
        End If
End Sub

Private Sub Text_BB_PROD_CD_LostFocus()
If Text_BB_PROD_CD.Text <> "" Then
   If (Len(Text_BB_PROD_CD.Text) < Text_BB_PROD_CD.MaxLength) Then
      Call Gp_MsgBoxDisplay("产品分类不符合规范！")
      'Text_PROD_CD.Text = ""
      Text_BB_PROD_CD.SetFocus
   End If
End If
End Sub



Private Sub Text_BB_REC_STS_Change()
If Not Text_BB_REC_STS.Text = "" Then
    If Not Text_BB_REC_STS.Text = "1" Then
      If Not Text_BB_REC_STS.Text = "2" Then
         If Not Text_BB_REC_STS.Text = "3" Then
            Call MsgBox("状态代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
            Text_BB_REC_STS.Text = ""

         End If
      End If
    End If
End If
If Text_BB_REC_STS.Text = "3" Then
   Check_ord_END.Value = 1
   Check_ord_END.Enabled = False
  Else
   Check_ord_END.Value = 0
   Check_ord_END.Enabled = True
End If
End Sub

Private Sub Text_BB_REC_STS_GotFocus()
'  If Check_ord_END.Value = 1 Then
'    Text_BB_REC_STS.Text = 3
'    Text_BB_REC_STS.Locked = True
'    Text_BB_REC_STS.BackColor = &H8080FF
'    Else
'    Text_BB_REC_STS.Locked = False
'    Text_BB_REC_STS.BackColor = &HFFFFFF
'  End If
End Sub

Private Sub Text_BB_REC_STS_KeyUp(KeyCode As Integer, Shift As Integer)
   Text_BB_REC_STS_Name = ""
   If KeyCode = vbKeyF4 Then
 
        DD.Switch = "MS"

        DD.sKey = "Z0005"


        DD.rControl.Add Item:=Text_BB_REC_STS
        DD.rControl.Add Item:=Text_BB_REC_STS_Name
   
        DD.nameType = "2"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
       
        
        Call Gf_Common_DD(M_CN1, KeyCode)
       

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_REC_STS.Text)) = Text_BB_REC_STS.MaxLength Then
       '  Gf_ComnNAME_Find( 连接字符串, DD.sKEy内容 ,DD.nameType)
       ' Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        Text_BB_REC_STS_Name.Text = Gf_ComnNameFind(M_CN1, "Z0005", Text_BB_REC_STS.Text, 2)
    Else
        Text_BB_REC_STS_Name.Text = ""
    End If
End Sub

Private Sub Text_BB_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Text_BB_PROD_CD_mate = ""
   If KeyCode = vbKeyF4 Then
 
        DD.Switch = "MS"

        DD.sKey = "B0005"


        DD.rControl.Add Item:=Text_BB_PROD_CD
        DD.rControl.Add Item:=Text_BB_PROD_CD_mate
   
        DD.nameType = "2"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
       
        
        Call Gf_Common_DD(M_CN1, KeyCode)
       

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_PROD_CD.Text)) = Text_BB_PROD_CD.MaxLength Then
       '  Gf_ComnNAME_Find( 连接字符串, DD.sKEy内容 ,DD.nameType)
       ' Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        Text_BB_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_BB_PROD_CD.Text, 2)
    Else
        Text_BB_PROD_CD_mate.Text = ""
    End If
End Sub
Private Sub Text_BB_DEST_CD_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then

            DD.Switch = "MS"
            DD.rControl.Add Item:=Text_BB_DEST_CD
            DD.rControl.Add Item:=Text_BB_DEST_CD_mate

            DD.nameType = "1"

            Call Gf_Destination_DD(M_CN1, KeyCode)

            Exit Sub

    End If

    If Len(Trim(Text_BB_DEST_CD)) = Text_BB_DEST_CD.MaxLength Then
        Text_BB_DEST_CD.Text = Gf_DestNameFind(M_CN1, Trim(Text_BB_DEST_CD.Text), 1)
    Else
        Text_BB_DEST_CD_mate.Text = ""
    End If
        
End Sub

Private Sub Text_ORD_ITEM_Change()

   If Text_ORD_ITEM.Text <> "" Then
        If Val(Text_ORD_ITEM.Text) > iCount Or Val(Text_ORD_ITEM.Text) < 0 Or Text_ORD_ITEM.Text = "00" Then
        Call MsgBox("订单序号输入不正确!" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
        Text_ORD_ITEM.Text = ""
        End If

  End If

End Sub


Private Sub Text_ORD_ITEM_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub Text_ORD_ITEM_LostFocus()
    Dim S As String
  
        If Len(Text_ORD_ITEM.Text) = 1 Then
         S = Text_ORD_ITEM.Text
         Text_ORD_ITEM.Text = "0" + S
         
        End If
  
End Sub

Private Sub VScroll1_Change()
VScroll1.Min = iCount

Select Case VScroll1.Value
Case 1 To 9
Text_ORD_ITEM.Text = "0" & VScroll1.Value
Case 10 To 99
Text_ORD_ITEM.Text = VScroll1.Value

End Select
End Sub

Private Function txt_KeyPress(KeyAscii As Integer) As Integer

        Select Case KeyAscii
               
               Case Is <= 32
                    txt_KeyPress = KeyAscii
               Case 48 To 57
                    txt_KeyPress = KeyAscii
'               Case 46
'                    txt_KeyPress = KeyAscii
               Case Else
                    txt_KeyPress = 0
        End Select
                    
End Function
