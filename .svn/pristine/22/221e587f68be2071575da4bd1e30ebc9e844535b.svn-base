VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form slab_confirm 
   Caption         =   "板坯跨区选择"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   6735
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton confirm_cmd 
      Caption         =   "确定"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4575
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      _Version        =   393216
      _ExtentX        =   10610
      _ExtentY        =   8070
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
      MaxCols         =   5
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "slab_confirm.frx":0000
   End
End
Attribute VB_Name = "slab_confirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kqnum As String
'Dim Mc1 As New Collection           'Master Collection


Private Sub Command1_Click()
kqnum = "&&"
Unload Me
End Sub

Private Sub confirm_cmd_Click()
Dim I As Integer


With ss1
    kqnum = ""
    For I = 1 To .MaxRows
    .Row = I
    .Col = 5
   If .Text = "1" Then
   .Col = 1
   kqnum = kqnum + .Text
   .Col = 2
   kqnum = kqnum + .Text
   .Col = 3
   kqnum = kqnum + .Text
   .Col = 4
   kqnum = kqnum + .Text
   kqnum = kqnum + "&&"
   End If
    Next
End With
Unload Me
End Sub



Private Function formstart()

Dim sQuery As String
sQuery = "select a.yard_type,a.zone_type,a.yard_row,a.yard_column from fp_stdyard a "
sQuery = sQuery + "where prod_type='S' "
sQuery = sQuery + "group by a.yard_type,"
sQuery = sQuery + "A.zone_type,a.yard_row,a.yard_column "
sQuery = sQuery + "order by a.yard_type,"
 sQuery = sQuery + "a.zone_type,a.yard_row,a.yard_column"
'If Gf_Sp_Display(M_CN1, ss1, squery, Sc1.Item("pColumn"), True) Then
On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If M_CN1 Is Nothing Then
        Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With ss1

        
        
        .ReDraw = False
        .MaxRows = 0: iCount = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        .MaxRows = UBound(ArrayRecords, 2) + 1
    
        For iRowCount = 0 To .MaxRows - 1
        
            .Row = iRowCount + 1
            
            For iColcount = 0 To 3
            
                .Col = iColcount + 1
                
     
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
               
                
            Next iColcount
            
        Next iRowCount
            
 

            'lControl Lock
            For iCount = 1 To 5

                .Protect = True
                .BlockMode = True: .Lock = True
                .BlockMode = False

            Next iCount

        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Call Gp_MsgBoxDisplay("Gf_Sp_Display Error : " & sQuery)
    Screen.MousePointer = vbDefault



End Function

Private Sub Form_Load()
Call formstart
End Sub
