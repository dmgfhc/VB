VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0020C 
   Caption         =   "����ָʾ��ϸ��ѯ - AQC0020C"
   ClientHeight    =   9090
   ClientLeft      =   165
   ClientTop       =   960
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_SMP_NO 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1590
      TabIndex        =   0
      Tag             =   "�������"
      Top             =   120
      Width           =   2325
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   210
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "�������"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   210
      Top             =   555
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "����"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   1530
      Top             =   555
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "¯��"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   2670
      Top             =   555
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   556
      Caption         =   "��׼��"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   4
      Left            =   4380
      Top             =   555
      Width           =   1950
      _ExtentX        =   3440
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   5
      Left            =   6330
      Top             =   555
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      Caption         =   "���к�"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   6
      Left            =   7680
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "������;"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   7
      Left            =   8910
      Top             =   555
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "�������"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   8
      Left            =   10110
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "�������"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   9
      Left            =   11340
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   10
      Left            =   12570
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "ȡ������"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   11
      Left            =   13800
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "ȡ������"
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
   Begin InDate.ULabel lbl_STLGRD 
      Height          =   345
      Left            =   210
      Top             =   855
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_HEAT_NO 
      Height          =   345
      Left            =   1530
      Top             =   855
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_STDSPEC 
      Height          =   345
      Left            =   2670
      Top             =   855
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_ORD_NO 
      Height          =   345
      Left            =   4380
      Top             =   855
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_ORD_ITEM 
      Height          =   345
      Left            =   6330
      Top             =   855
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_ENDUSE_CD 
      Height          =   345
      Left            =   7680
      Top             =   855
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_ORD_THK 
      Height          =   345
      Left            =   8910
      Top             =   855
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   ""
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_ORD_WID 
      Height          =   345
      Left            =   10110
      Top             =   855
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_CUST_CD 
      Height          =   345
      Left            =   11340
      Top             =   855
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_SMP_CNT 
      Height          =   345
      Left            =   12570
      Top             =   855
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel lbl_SMP_LEN 
      Height          =   345
      Left            =   13800
      Top             =   855
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7950
      Left            =   180
      TabIndex        =   1
      Top             =   1260
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   14023
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AQC0020C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   7950
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   8430
         _Version        =   393216
         _ExtentX        =   14870
         _ExtentY        =   14023
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0020C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   7950
         Left            =   8520
         TabIndex        =   3
         Top             =   0
         Width           =   6375
         _Version        =   393216
         _ExtentX        =   11245
         _ExtentY        =   14023
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
         MaxCols         =   3
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0020C.frx":0401
      End
   End
End
Attribute VB_Name = "AQC0020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       ��������
'-- Sub_System Name   �ж�����
'-- Program Name      ����ָʾ��ϸ��ѯ
'-- Program ID        AQC0020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.8.18
'-- Description       ����ָʾ��ϸ��ѯ
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


Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim arrChem(4, 61) As String

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
    
            
    Call Gp_Sp_Setting(ss1, False)
    Call Gp_Sp_Setting(SS2, False)
    
    Call Gp_Sp_ReadOnlySet(ss1)
    Call Gp_Sp_ReadOnlySet(SS2)
   
    Call MDIMain.FormMenuSetting(Me, "Refer", "FS", sAuthority)

    Call subFormClear
    
    Call Gp_Sp_ColGet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(SS2, "Q-System.INI", Me.Name)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

    Screen.MousePointer = vbDefault
            
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(SS2, "Q-System.INI", Me.Name)
            
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    Call subFormClear
                    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            
End Sub


Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim arr As Variant
    Dim V_ORD_NO As String
    Dim V_ORD_ITEM As String
    
        
'On Error GoTo Error_Rtn
        
    If Trim(txt_SMP_NO.Text) = "" Then
        sMesg = "������� �������룡"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
        
    ss1.MaxRows = 0
    SS2.MaxRows = 0
        
    Set AdoRs = New adodb.Recordset
    
    sQuery = "{call AQC0020C.P_REFER('" + Trim(txt_SMP_NO.Text) + "')}"
                    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.EOF Or AdoRs.BOF Then GoTo Error_Rtn
                       
    Call subSetTitle
                       
    ArrayRecords = AdoRs.GetRows
    AdoRs.Close
    
    Call subLoadMaster(ArrayRecords)
    
    If VarType(ArrayRecords(4, 0)) = vbNull Or VarType(ArrayRecords(5, 0)) = vbNull Then
       sMesg = "������/�������к�Ϊ��"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    V_ORD_NO = ArrayRecords(4, 0)
    V_ORD_ITEM = ArrayRecords(5, 0)
            
    sQuery = "{call AQC0020C.P_REFER_SS1('" + ArrayRecords(4, 0) + "','" + ArrayRecords(5, 0) + "')}"
    
    Erase ArrayRecords
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Set AdoRs = M_CN1.Execute(sQuery)
    
    If AdoRs.EOF Or AdoRs.BOF Then GoTo Error_Rtn
            
    ArrayRecords = AdoRs.GetRows
    AdoRs.Close
    
    Call subSetDecCd(ArrayRecords)
    
    Erase ArrayRecords
    
    Call subSpreadView1
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    
    '--------------------���û���Ŀ��ʾ  ����  2012.11.20-----------------------------------------------------
    
    Erase ArrayRecords

    sQuery = "{call AQC0020C.P_SREFER_CONFIG('" + V_ORD_NO + "','" + V_ORD_ITEM + "')}"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Set AdoRs = M_CN1.Execute(sQuery)
    
    If Not AdoRs.EOF And Not AdoRs.BOF Then
      ArrayRecords = AdoRs.GetRows
      Call subSpreadView_Config(ArrayRecords)
    End If
    
    AdoRs.Close
    Erase ArrayRecords
    
    '-----------------------------------------------------------------------------------------------------------
    
    sQuery = "{call AQC0020C.P_REFER_SS2('" + Trim(txt_SMP_NO.Text) + "')}"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.EOF Or AdoRs.BOF Then GoTo Error_Rtn
    ArrayRecords = AdoRs.GetRows
    AdoRs.Close
    
    Call subSpreadView2(ArrayRecords)
    Call Gp_Sp_EvenRowBackcolor(SS2)
    
    Exit Sub
    
Error_Rtn:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
 
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

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
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

Private Sub subFormClear()
    ss1.MaxRows = 0
    SS2.MaxRows = 0
    
    txt_SMP_NO.Text = ""
    lbl_STLGRD.Caption = ""
    lbl_HEAT_NO.Caption = ""
    lbl_STDSPEC.Caption = ""
    lbl_ORD_NO.Caption = ""
    lbl_ORD_ITEM.Caption = ""
    lbl_ENDUSE_CD.Caption = ""
    lbl_ORD_THK.Caption = ""
    lbl_ORD_WID.Caption = ""
    lbl_CUST_CD.Caption = ""
    lbl_SMP_CNT.Caption = ""
    lbl_SMP_LEN.Caption = ""
    
End Sub

Private Sub subLoadMaster(ByVal arr As Variant)
    
    txt_SMP_NO.Text = GF_NullChange(arr(0, 0))
    lbl_STLGRD.Caption = GF_NullChange(arr(1, 0))
    lbl_HEAT_NO.Caption = GF_NullChange(arr(2, 0))
    lbl_STDSPEC.Caption = GF_NullChange(arr(3, 0))
    lbl_ORD_NO.Caption = GF_NullChange(arr(4, 0))
    lbl_ORD_ITEM.Caption = GF_NullChange(arr(5, 0))
    lbl_ENDUSE_CD.Caption = GF_NullChange(arr(6, 0))
    lbl_ORD_THK.Caption = GF_NullChange(arr(7, 0))
    lbl_ORD_WID.Caption = GF_NullChange(arr(8, 0))
    lbl_CUST_CD.Caption = GF_NullChange(arr(9, 0))
    lbl_SMP_CNT.Caption = GF_NullChange(arr(10, 0))
    lbl_SMP_LEN.Caption = GF_NullChange(arr(11, 0))
    
End Sub

Private Sub subSetTitle()
    
    arrChem(0, 0) = "����������"
    arrChem(0, 1) = "����ǿ������"
    arrChem(0, 2) = "�����쳤������"
    arrChem(0, 3) = "����������"
    arrChem(0, 4) = "������ȷ����������������"
    arrChem(0, 5) = "�涨�Ǳ����쳤Ӧ������"
    arrChem(0, 6) = "�涨���쳤Ӧ������"
    arrChem(0, 7) = "�涨�����쳤Ӧ������"
'20090803 SUN BIN STARTR
    arrChem(0, 8) = "׷������������"
    arrChem(0, 9) = "׷�ӿ���ǿ������"
    arrChem(0, 10) = "׷�Ӷ����쳤������"
    arrChem(0, 11) = "׷������������"
    arrChem(0, 12) = "׷�ӹ涨�Ǳ����쳤Ӧ������"
    arrChem(0, 13) = "׷�ӹ涨���쳤Ӧ������"
    arrChem(0, 14) = "׷�ӹ涨�����쳤Ӧ������"
'20090803 SUN BIN END
    arrChem(0, 15) = "�ߜ�����������"
    arrChem(0, 16) = "�ߜؿ���ǿ������"
    arrChem(0, 17) = "�ߜض����쳤������"
    arrChem(0, 18) = "�ߜ�����������"
     arrChem(0, 19) = "����������ȷ����������������"
    arrChem(0, 20) = "�ߜع涨�Ǳ����쳤Ӧ������"
    arrChem(0, 21) = "�ߜع涨�����쳤Ӧ������"
'20090803 SUN BIN START
    arrChem(0, 22) = "׷�Ӹߜ�����������"
    arrChem(0, 23) = "׷�Ӹߜؿ���ǿ������"
    arrChem(0, 24) = "׷�Ӹߜض����쳤������"
    arrChem(0, 25) = "׷�Ӹߜ�����������"
    arrChem(0, 26) = "׷�Ӹߜع涨�Ǳ����쳤Ӧ������"
    arrChem(0, 27) = "׷�Ӹߜع涨�����쳤Ӧ������"
'20090803 SUN BIN END
    arrChem(0, 28) = "�������"
    arrChem(0, 29) = "׷�ӳ������"
    arrChem(0, 30) = "ʱЧ�������"
    arrChem(0, 31) = "׷��ʱЧ�������"
    arrChem(0, 32) = "��������"
'20090803 SUN BIN START
    arrChem(0, 33) = "׷����������"
'20090803 SUN BIN END
    arrChem(0, 34) = "Ӳ������"
'20090803 SUN BIN START
    arrChem(0, 35) = "׷��Ӳ������"
'20090803 SUN BIN END
    arrChem(0, 36) = "UST����"
    arrChem(0, 37) = "��̼������"
    arrChem(0, 38) = "����������"
    arrChem(0, 39) = "��ӡ����"
    arrChem(0, 40) = "�Ͽ�����"
    arrChem(0, 41) = "�ǽ�����������"
    arrChem(0, 42) = "�������"
    arrChem(0, 43) = "��ƽ����"
    arrChem(0, 44) = "��͸������"
    arrChem(0, 45) = "������������"
    arrChem(0, 46) = "����Ӳ������"
    arrChem(0, 47) = "������������"
    arrChem(0, 48) = "���︯ʴ����"
    arrChem(0, 49) = "���︯ʴ����"
    arrChem(0, 50) = "����˺������"
    arrChem(0, 51) = "��״��֯"
    arrChem(0, 52) = "�����徧����"
    'louyannan 20101119 start
    
   
    arrChem(0, 53) = "ʣ������"
    arrChem(0, 54) = "NDT����˺������"
    
    'louyannan 20101119 end
    
    'edit by gengxueyu 20110211 start
    arrChem(0, 55) = "���ȱ����쳤��UEL"
    arrChem(0, 56) = "׷�Ӿ��ȱ����쳤��UEL"
    arrChem(0, 57) = "׷��Ӧ����������Ŀ1"
    arrChem(0, 58) = "׷��Ӧ����������Ŀ2"
    arrChem(0, 59) = "׷��Ӧ����������Ŀ3"
    arrChem(0, 60) = "׷��Ӧ����������Ŀ4"
    arrChem(0, 61) = "׷��Ӧ����������Ŀ5"
    'edit by gengxueyu 20110211 end
    
    
    
End Sub

Private Sub subSetDecCd(ByVal strArr As Variant)
 'gengxueyu 20110211 start
 Dim i As Integer
    
    If UBound(strArr) < 63 Then Exit Sub
    
    For i = 0 To 61
        
        arrChem(1, i) = NullCheck(strArr(i, 0), "")
    
    Next i
    
    For i = 0 To 61
        
        arrChem(2, i) = NullCheck(strArr(i + 62, 0))
    
    Next i
    
    For i = 0 To 61
        
        arrChem(3, i) = NullCheck(strArr(i + 124, 0))
    
    Next i
    
    For i = 0 To 34
        
        arrChem(4, i) = NullCheck(strArr(i + 186, 0))
    
    Next i
    
    'gengxueyu 20110211 end
        
End Sub

Private Sub subSpreadView1()

    Dim i As Integer
    
    With ss1
        
        .MaxRows = 62
        
        For i = 1 To 62
            .Row = i: .Col = 1
            .Text = arrChem(0, i - 1)
        Next i
    
        For i = 1 To 62
            .Row = i: .Col = 2
            .Text = arrChem(1, i - 1)
        Next i
        
        For i = 1 To 62
            .Row = i: .Col = 3
            If i >= 2 And i <= 7 Then
                 If arrChem(1, i - 1) <> "" Then
                    .Text = arrChem(2, i - 1)
                 Else
                    .Text = ""
                 End If
            ElseIf i >= 2 And i <= 7 Then
                 If arrChem(1, i - 1) <> "" Then
                    .Text = arrChem(2, i - 1)
                 Else
                    .Text = ""
                 End If
            ElseIf i >= 10 And i <= 13 Then
                 If arrChem(1, i - 1) <> "" Then
                    .Text = arrChem(2, i - 1)
                 Else
                    .Text = ""
                 End If
            
            Else
                .Text = arrChem(2, i - 1)
            End If
            
        Next i
        
        For i = 1 To 62
            .Row = i: .Col = 4
            .Text = arrChem(3, i - 1)
        Next i
        
        For i = 1 To 35
            .Row = i: .Col = 5
            .Text = arrChem(4, i - 1)
        Next i
            
    End With
    
    Call subSpreadCheck1

End Sub

Private Sub subSpreadCheck1()
    
 Dim i As Long
 Dim j As Long
    
    With ss1
        
        For i = 1 To .MaxRows
                                    
            If Gf_Get_Cell_Value(ss1, i, 2) = "" And Gf_Get_Cell_Value(ss1, i, 3) = "" Then
                .Row = i
                .RowHidden = True
            Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j
            End If
        Next i
                
    End With
   
    
End Sub

Private Sub subSpreadView2(ByVal strArr As Variant)

    Dim i As Integer
    
    If UBound(strArr, 2) < 0 Then Exit Sub
    
    With SS2
        
        .MaxRows = UBound(strArr, 2) + 1
        
        For i = 1 To UBound(strArr, 2) + 1
            .Row = i
            .Col = 1: .Text = GF_NullChange(strArr(0, i - 1))
            .Col = 2: .Text = GF_NullChange(strArr(1, i - 1))
            .Col = 3: .Text = GF_NullChange(strArr(2, i - 1))
        
        Next i
    
        
            
    End With
    
  '  Call subSpreadCheck1

End Sub

Private Sub ss2_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss2, NewRow)
End Sub

Private Sub subSpreadView_Config(ByVal strArr As Variant)

    Dim i As Integer
    Dim OLD_MAXROWS As Integer
    
    If UBound(strArr, 2) < 0 Then Exit Sub
    
    With ss1
        OLD_MAXROWS = .MaxRows
        .MaxRows = .MaxRows + UBound(strArr, 2) + 1

        For i = 1 To UBound(strArr, 2) + 1
            .Row = OLD_MAXROWS + i
            .Col = 1: .Text = GF_NullChange(strArr(0, i - 1))
            .Col = 2: .Text = GF_NullChange(strArr(1, i - 1))
            .Col = 3: .Text = GF_NullChange(strArr(2, i - 1))
            .Col = 4: .Text = GF_NullChange(strArr(3, i - 1))
            .Col = 5: .Text = GF_NullChange(strArr(4, i - 1))
        Next i
            
    End With
    
   'subSpreadCheck1 ����������Ŀ���ж�����Ϊ�յ��У�����д��һ��˳���
    Call subSpreadCheck1

End Sub
