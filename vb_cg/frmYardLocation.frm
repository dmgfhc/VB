VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmYardLocation 
   Caption         =   "Plate Yard Location Selection"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLocation 
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
      ItemData        =   "frmYardLocation.frx":0000
      Left            =   1125
      List            =   "frmYardLocation.frx":0002
      TabIndex        =   6
      Tag             =   "高炉代码"
      Top             =   135
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   5760
      Width           =   1080
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确认"
      Height          =   375
      Left            =   4260
      TabIndex        =   0
      Top             =   5760
      Width           =   1080
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   5115
      Left            =   75
      TabIndex        =   2
      Top             =   540
      Width           =   1980
      _Version        =   393216
      _ExtentX        =   3493
      _ExtentY        =   9022
      _StockProps     =   64
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
      MaxCols         =   1
      MaxRows         =   20
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmYardLocation.frx":0004
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   5115
      Left            =   2115
      TabIndex        =   3
      Top             =   540
      Width           =   4365
      _Version        =   393216
      _ExtentX        =   7699
      _ExtentY        =   9022
      _StockProps     =   64
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
      MaxCols         =   3
      MaxRows         =   30
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmYardLocation.frx":03AB
   End
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   135
      Top             =   135
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "Yard"
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
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   930
      TabIndex        =   5
      Top             =   5880
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "垛位号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   150
      TabIndex        =   4
      Top             =   5880
      Width           =   810
   End
End
Attribute VB_Name = "frmYardLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL         As String
Dim iDR         As Long

Private Sub cboLocation_Click()
    Call SerchLocation
End Sub

Private Sub Form_Activate()
    ss1.ReDraw = False
    ss2.ReDraw = False
    
    Call Gp_Sp_Setting(ss1)
    Call Gp_Sp_Setting(ss2)
    
    Call ComboList
    Call SerchLocation
    ss1.ReDraw = True
    ss2.ReDraw = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'    If Len(Trim(Label2.Caption)) > 0 Then AGZ3010C.Tag = Label2.Caption
'    Unload Me
End Sub

Private Sub SerchLocation()
    Set AdoRs = New ADODB.Recordset

    SQL = " SELECT      YARD_ADDR                 " & vbCrLf
    SQL = SQL & "  FROM GP_PLATEYARD              " & vbCrLf
    SQL = SQL & " WHERE SUBSTR(YARD_ADDR,1,2) =  '" & cboLocation.Text & "'" & vbCrLf
    SQL = SQL & " GROUP BY YARD_ADDR              " & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    ss1.MaxRows = AdoRs.RecordCount
    
    iDR = 1
    Do Until AdoRs.EOF
        ss1.Row = iDR
        ss1.Col = 1
        ss1.Text = AdoRs.Fields("YARD_ADDR")
            
        AdoRs.MoveNext
        iDR = iDR + 1
    Loop
    
    AdoRs.Close
    
    ss1.BlockMode = True:  ss1.Row = -1:   ss1.Col = -1:  ss1.Lock = True:   ss1.BlockMode = False
    ss2.BlockMode = True:  ss2.Row = -1:   ss2.Col = -1:  ss2.Lock = True:   ss2.BlockMode = False
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sYardLoc    As String
    
    If Row < 1 Then Exit Sub

    ss1.Col = Col
    ss1.Row = Row
    sYardLoc = Trim(ss1.Text)
    
    Set AdoRs = New ADODB.Recordset
    
    Label2.Caption = ""
    
'    SQL = " SELECT      YARD_ADDR||Lpad(BED_SEQ, 3, '0')  YARD_ADDR" & vbCrLf
    SQL = " SELECT      YARD_ADDR                           " & vbCrLf
    SQL = SQL & "      ,BED_SEQ                             " & vbCrLf
    SQL = SQL & "      ,NVL(PLATE_NO,' ')          PLATE_NO " & vbCrLf
    SQL = SQL & "  FROM GP_PLATEYARD                        " & vbCrLf
    SQL = SQL & " WHERE YARD_ADDR  =  '" & sYardLoc & "'    " & vbCrLf
    SQL = SQL & " ORDER BY YARD_ADDR, BED_SEQ               " & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    ss2.MaxRows = AdoRs.RecordCount
    
    iDR = 1
    Do Until AdoRs.EOF
        ss2.Row = iDR
        ss2.Col = 1
        ss2.Text = AdoRs.Fields("YARD_ADDR")
        ss2.Col = 2
        ss2.Text = Val(AdoRs.Fields("BED_SEQ") & "")
        ss2.Col = 3
        ss2.Text = AdoRs.Fields("PLATE_NO")
            
        AdoRs.MoveNext
        iDR = iDR + 1
    Loop
    
    AdoRs.Close
    
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub

    ss2.Col = 1: ss2.Row = Row
    Label2.Caption = Trim(ss2.Text)
    ss2.Col = 2: Label2.Caption = Label2.Caption & Trim(ss2.Text)
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    ss2.Col = 1: ss2.Row = Row
    Label2.Caption = Trim(ss2.Text)
    ss2.Col = 2: Label2.Caption = Label2.Caption & Trim(ss2.Text)
    
'    If Len(Trim(Label2.Caption)) > 0 Then AGZ3010C.Tag = Trim(Label2.Caption)
'    Unload Me
End Sub

Private Sub ComboList()
    Set AdoRs = New ADODB.Recordset

    SQL = " SELECT      SUBSTR(YARD_ADDR,1,2) ADDR    " & vbCrLf
    SQL = SQL & "  FROM GP_PLATEYARD                  " & vbCrLf
    SQL = SQL & " GROUP BY SUBSTR(YARD_ADDR,1,2)      " & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    cboLocation.Clear

    Do Until AdoRs.EOF
        cboLocation.AddItem Trim(AdoRs.Fields("ADDR").Value)
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
    cboLocation.ListIndex = 0
End Sub
