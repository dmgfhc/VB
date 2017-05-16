VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB4122C 
   Caption         =   "钢板转坯料履历查询_ACB4122C"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   1296
      _Version        =   196610
      Begin VB.TextBox TXT_MAT_NO 
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
         Left            =   6390
         MaxLength       =   14
         TabIndex        =   3
         Tag             =   "物料号"
         Top             =   240
         Width           =   1485
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   360
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "转换日期"
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
      Begin InDate.UDate DTP_PROD_FR 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Tag             =   "INS_DATE"
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
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
      Begin InDate.UDate DTP_PROD_TO 
         Height          =   315
         Left            =   3465
         TabIndex        =   2
         Tag             =   "INS_DATE"
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   5160
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "物料号"
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
         ForeColor       =   0
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9915
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   17489
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "已转坯料钢板信息"
      TabPicture(0)   =   "ACB4122C.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ss1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "新生成坯料钢板信息"
      TabPicture(1)   =   "ACB4122C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SS2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin FPSpread.vaSpread ss1 
         Height          =   7155
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   12621
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   106
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB4122C.frx":0038
      End
      Begin FPSpread.vaSpread SS2 
         Height          =   7935
         Left            =   -75000
         TabIndex        =   6
         Top             =   480
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   13996
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   104
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB4122C.frx":2B99
      End
   End
End
Attribute VB_Name = "ACB4122C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public STR1 As String
Public BASE As String
Public AIMNO As String
Public Refer_Fl As String
Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Dim sQuery As String
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
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer

Const SPD_WGT = 17
Const SPD_DEL_TO_DATE = 6
Const SS1_PLATE_NO = 1
Const SS1_ORD_NO = 32   '20150331 在25列增加了综判时间  29--30->32
Const SS1_ORD_ITEM = 33 '20150331 在25列增加了综判时间  30--31->33
Const SS1_URGNT_FL = 82          '紧急订单绿色标记  2012-11-07 by CaoLei    65->67    '20150331 在25列增加了综判时间  76--77->79
Const SS1_RH_FL = 31          '是否走真空      '20150331 在25列增加了综判时间  28--29->31
Const SS1_KEY_ORD_FL = 84         '重点合同   67->69     '20150331 在25列增加了综判时间  78--79 zhouyan   79->81
 

Private Sub Form_Define()

    Dim iRow As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         
          Call Gp_Ms_Collection(DTP_PROD_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(DTP_PROD_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
    
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    
    For iRow = 5 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4122C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
    Call Gp_Sp_Collection(SS2, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(SS2, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(SS2, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(SS2, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    
    For iRow = 5 To ss1.MaxCols
        Call Gp_Sp_Collection(SS2, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
    
    'Spread_Collection
    sc2.Add Item:=SS2, Key:="Spread"
    sc2.Add Item:="ACB4122C.P_SREFER1", Key:="P-R"
    sc2.Add Item:=pColumn1, Key:="pColumn"
    sc2.Add Item:=nColumn1, Key:="nColumn"
    sc2.Add Item:=aColumn1, Key:="aColumn"
    sc2.Add Item:=mColumn1, Key:="mColumn"
    sc2.Add Item:=iColumn1, Key:="iColumn"
    sc2.Add Item:=lColumn1, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=SS2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
'    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols - 1, True)
  ' Call Gp_Sp_ColHidden(ss1, 2, True)

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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    

   
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
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
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub



Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    
End Sub

Public Sub Form_Ref()

  
    Dim iRow As Long
    Dim iSumWgt As Double
 
  Select Case SSTab1.Tab
    
           Case 0
                If Gf_Sp_Refer(M_CN1, Proc_Sc("SC"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

                   
                  
                   
                   ss1.OperationMode = OperationModeNormal
                End If
           
           Case 1
               If Gf_Sp_Refer(M_CN1, Proc_Sc("SC2"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)


                   SS2.OperationMode = OperationModeNormal
                End If
 End Select
 
 With ss1

         .MaxRows = ss1.MaxRows + 1
          .ROW = .MaxRows:         .Col = 1
          .Text = "合计"
            For iRow = 1 To .MaxRows - 1
                .ROW = iRow
                .Col = SPD_WGT
                iSumWgt = iSumWgt + Val(.Text)
            Next iRow
          .ROW = ss1.MaxRows:          .Col = SPD_WGT
           iSumWgt = Round(iSumWgt, 3)
          .Text = iSumWgt

   End With

End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

End Sub



Public Sub Form_Exc()
    
     Select Case SSTab1.Tab
    
           Case 0
    
               Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
               
           Case 1
    
               Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
    End Select


End Sub

Public Sub Form_Exit()
    Unload Me
End Sub



Private Sub SSTab1_Click(Previoustab As Integer)

      Select Case SSTab1.Tab
    
           Case 0
               
                ULabel1.Caption = "转换日期"
           
          Case 1
              
                  
                 ULabel1.Caption = "生产日期"
           
 End Select
    



End Sub
