VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form OrderCopy 
   Caption         =   "订单复制"
   ClientHeight    =   3690
   ClientLeft      =   2805
   ClientTop       =   3435
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   11775
   Begin VB.TextBox TXT_EMP_ID 
      Height          =   405
      Left            =   4725
      TabIndex        =   5
      Top             =   2265
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txt_cust_cd 
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
      Left            =   5265
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "客户代码"
      Top             =   1170
      Width           =   960
   End
   Begin VB.TextBox txt_cust_cd_name 
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
      Left            =   5265
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "end_cust_cd_name"
      Top             =   1530
      Width           =   4425
   End
   Begin Threed.SSCommand cmdcopy 
      Height          =   420
      Left            =   180
      TabIndex        =   2
      Top             =   45
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   741
      _Version        =   196609
      Caption         =   "复制"
   End
   Begin Threed.SSCommand cmdexit 
      Height          =   420
      Left            =   1485
      TabIndex        =   1
      Top             =   45
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   741
      _Version        =   196609
      Caption         =   "退出"
   End
   Begin InDate.ULabel ord_no_to 
      Height          =   315
      Left            =   5265
      Top             =   675
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   0
      BackColor       =   -2147483639
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
   Begin InDate.ULabel ord_no_fr 
      Height          =   315
      Left            =   1530
      Top             =   675
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   0
      BackColor       =   -2147483639
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
   Begin FPSpread.vaSpread ss1 
      Height          =   2265
      Left            =   180
      TabIndex        =   0
      Top             =   1170
      Width           =   2895
      _Version        =   393216
      _ExtentX        =   5106
      _ExtentY        =   3995
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
      MaxCols         =   2
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "OrderCopy.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   675
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "订单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3915
      Top             =   675
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "订单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Left            =   3915
      Top             =   1170
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "客户"
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9810
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3240
      Picture         =   "OrderCopy.frx":0309
      Top             =   1170
      Width           =   480
   End
End
Attribute VB_Name = "OrderCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim sc1 As New Collection           'Spread Collection

Private Sub cmdcopy_Click()
Call txt_cust_cd_LostFocus
Call Gp_Process_Exec

End Sub

Private Sub cmdexit_Click()

   Unload Me

End Sub

Private Sub Form_Load()

    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    sc1.Add Item:=ss1, Key:="Spread"
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    ss1.OperationMode = OperationModeNormal
   
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing

    Set sc1 = Nothing
    
End Sub


Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_cd_name.Text = ""
    End If

End Sub
Private Sub txt_cust_cd_LostFocus()

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    If Len(txt_cust_cd.Text) <> 0 Then
    
        sQuery = "select * from bp_cust_cd  where cust_cd='" + txt_cust_cd.Text + "'"
        Set AdoRs = New ADODB.Recordset
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        If Not AdoRs.BOF And Not AdoRs.EOF Then
        
'            If VarType(adoRs.Fields("dome_fl").Value) = vbNull Then
'    '         If AdoRs.BOF = True Or AdoRs.EOF = True Then
'               txt_dome_fl.Text = ""
'               txt_dome_fl_name.Text = ""
'            Else
'               txt_dome_fl.Text = adoRs.Fields("dome_fl").Value
'               Call txt_dome_fl_KeyUp(0, 0)
'            End If
            AdoRs.Close
            Set AdoRs = Nothing
            
       Else
       
            Call Gp_MsgBoxDisplay("客户代码不存在.......")
            
            txt_cust_cd.Text = ""
            txt_cust_cd.SetFocus
            
       End If
       
    End If
    
End Sub
Private Sub txt_cust_cd_Validate(Cancel As Boolean)
'
'  Dim sQuery As String
'  'Order_No Make
'  sQuery = "{call ABA1010C.P_ORD_NO ( '" + txt_cust_cd.Text + "' )}"
'  ord_no_to.Caption = Gf_CodeFind(M_CN1, sQuery)


End Sub

Public Sub Gp_Process_Exec()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim iRow, iCnt As Integer
    Dim sOrdNo_fr, sOrdNo_to As String
    Dim sOrdItems As String
    Dim sQuery As String
    
    If ss1.MaxRows < 1 Then
       Exit Sub
    End If
    
    ss1.Col = 1
    iCnt = 0
    sOrdItems = ""
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        If ss1.Text = "1" Then
           sOrdItems = sOrdItems & "1"
           iCnt = iCnt + 1
        Else
           sOrdItems = sOrdItems & "0"
        End If
    Next iRow
    
    If iCnt = 0 Then
       Call Gp_MsgBoxDisplay("没有选定订单序列.......")
       Exit Sub
    End If
    
    If txt_cust_cd.Text = "" Then
       Call Gp_MsgBoxDisplay("必须输入客户代码.......")
       Exit Sub
    End If
    
    
   'Order_No Make
   sQuery = "{call ABA1010C.P_ORD_NO ( '" + txt_cust_cd.Text + "' )}"
   ord_no_to.Caption = Gf_CodeFind(M_CN1, sQuery)

    
    Dim adoCmd As ADODB.Command

    Screen.MousePointer = vbHourglass

    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1

    sQuery = "{call ABA1000P ('" + ord_no_fr.Caption + "','" + sOrdItems + "','" + ord_no_to.Caption + "','" + TXT_EMP_ID.Text + "',?)}"
    
'    Debug.Print sQuery
    

        adoCmd.CommandText = sQuery

        adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))

        adoCmd.Execute , , adExecuteNoRecords

        'Process Error Check
        If adoCmd("arg_e_msg") <> "" Then
            ret_Result_ErrMsg = adoCmd("arg_e_msg")
            sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
            Call Gp_MsgBoxDisplay(sErrMessg)
            Set adoCmd = Nothing
            Exit Sub
        End If


    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub


