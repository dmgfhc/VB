VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AHC0030C 
   Caption         =   "提货单 - REPORT"
   ClientHeight    =   8775
   ClientLeft      =   4185
   ClientTop       =   2280
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   9090
   Begin VB.CommandButton Command1 
      Caption         =   "打印计量单"
      Height          =   420
      Left            =   7380
      TabIndex        =   4
      Top             =   255
      Visible         =   0   'False
      Width           =   1275
   End
   Begin Threed.SSCommand cmdprint 
      Height          =   420
      Left            =   3420
      TabIndex        =   2
      Top             =   270
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      _Version        =   196609
      Caption         =   "打印发货清单"
   End
   Begin Threed.SSCommand cmdexit 
      Height          =   420
      Left            =   5760
      TabIndex        =   1
      Top             =   285
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      _Version        =   196609
      Caption         =   "退出"
   End
   Begin InDate.ULabel LBL_TRNS_NO 
      Height          =   315
      Left            =   1530
      Tag             =   "发放提货单号"
      Top             =   585
      Width           =   1320
      _ExtentX        =   2328
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
   Begin InDate.ULabel LBL_PROD_CD 
      Height          =   315
      Left            =   4860
      Tag             =   "产品"
      Top             =   135
      Visible         =   0   'False
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   0
      BackColor       =   -2147483639
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel LBL_SHP_IST_NO 
      Height          =   315
      Left            =   1530
      Tag             =   "发货指示号"
      Top             =   135
      Width           =   1320
      _ExtentX        =   2328
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
      Height          =   7650
      Left            =   165
      TabIndex        =   0
      Top             =   990
      Width           =   8910
      _Version        =   393216
      _ExtentX        =   15716
      _ExtentY        =   13494
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
      SpreadDesigner  =   "AHC0030C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   135
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "提单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   3510
      Top             =   135
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "产品"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   180
      Top             =   585
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "发货明细号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBL_STDSPEC 
      BackColor       =   &H80000009&
      Height          =   435
      Left            =   6255
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1605
   End
End
Attribute VB_Name = "AHC0030C"
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
'---------------------------------------------------------------------------------------------
'------------------------------ Report Variable ----------------------------------------------
'---------------------------------------------------------------------------------------------
Dim xlApp       As Object
Dim xlSheet     As Object

Dim arrRecords1 As Variant      'sQueryHeadC
Dim arrRecords2 As Variant      'sQueryDetailC
Dim arrRecords3 As Variant      'sQueryDetailC

Dim sQuery      As String
Dim sErrMsg     As String
Dim sDate       As String
Dim adoRs       As ADODB.Recordset
Private Sub cmdexit_Click()

   Unload Me

End Sub

'-----------------------------------------------------------------------
'---------------------------- Report Main ------------------------------
'-----------------------------------------------------------------------
Private Sub cmdPrint_Click()
    
    If Trim(LBL_TRNS_NO.Caption) = "" Or Trim(LBL_PROD_CD.Caption) = "" Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    If LBL_PROD_CD.Caption = "PP" And _
       (Left(Trim(LBL_STDSPEC.Caption), 3) = "CCS" Or _
        Left(Trim(LBL_STDSPEC.Caption), 3) = "ABS" Or _
        Left(Trim(LBL_STDSPEC.Caption), 3) = "DNV" Or _
        Left(Trim(LBL_STDSPEC.Caption), 3) = "KST" Or _
        Left(Trim(LBL_STDSPEC.Caption), 2) = "LR" Or _
        Left(Trim(LBL_STDSPEC.Caption), 2) = "KR" Or _
        Left(Trim(LBL_STDSPEC.Caption), 2) = "NK" Or _
        Left(Trim(LBL_STDSPEC.Caption), 2) = "GL" Or _
        Left(Trim(LBL_STDSPEC.Caption), 2) = "BV" Or _
        Left(Trim(LBL_STDSPEC.Caption), 4) = "RINA") Then
         If subGetOracleData_1 = False Then
            Exit Sub
         End If
    Else
         If subGetOracleData = False Then
            Exit Sub
         End If
    End If
            
'    Call MillSheetPrint_C
    Screen.MousePointer = vbDefault
    
End Sub

'----------------------------- Oracle Data Select (To MDB ) -------------------------------
'---------------------------------------船板---------------------------------------------

Private Function subGetOracleData_1() As Boolean
    
'    Dim sQuery As String
'    Dim AdoRs As adodb.Recordset
'    Dim arrRecords1 As Variant      'sQueryHead
'    Dim arrRecords2 As Variant      'sQueryDetail
    
    On Error GoTo Err_Track
                
    Set adoRs = New ADODB.Recordset
    
    sQuery = "{call AHC0031C.P_REFER_HEAD("
    sQuery = sQuery + "'" + LBL_SHP_IST_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_TRNS_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_PROD_CD.Caption + "')}"

'-----------------------------------------------------------------------------
        
    adoRs.Open sQuery, M_CN1, adOpenKeyset

    If adoRs.EOF Then GoTo Err_Track
       
    arrRecords1 = adoRs.GetRows
    adoRs.Close
    
    sQuery = "{call AHC0031C.P_REFER_DETAIL("
    sQuery = sQuery + "'" + LBL_SHP_IST_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_TRNS_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_PROD_CD.Caption + "')}"
                                
    adoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If adoRs.EOF Then GoTo Err_Track
    
    arrRecords2 = adoRs.GetRows
    adoRs.Close
    Set adoRs = Nothing
    
'    Call subMdbUpdate(arrRecords1, arrRecords2)
    Call MillSheetPrint_C_1
    subGetOracleData_1 = True
    
    Exit Function
    
Err_Track:
        
    If IsObject(adoRs) = True Then
        Set adoRs = Nothing
    End If
        
End Function

Private Sub MillSheetPrint_C_1()
    Dim RowCnt      As Long
    Dim PrtCnt      As Long
    Dim LneCnt      As Long
    Dim pAry()      As String
    Dim sRow        As String
    Dim sRow2       As String
    
    If IsEmpty(arrRecords1) Or IsEmpty(arrRecords2) Then Exit Sub
    
    RowCnt = UBound(arrRecords2, 2)
    
    PrtCnt = -1
    LneCnt = 0
    
    ReDim pAry(1 To 25, 1 To 10)
    
    Do

        LneCnt = LneCnt + 1
        PrtCnt = PrtCnt + 1
        
'        pAry(LneCnt, 1) = PrtCnt + 1                                   ' 序号

        pAry(LneCnt, 1) = arrRecords2(0, PrtCnt) & ""                  ' 产品号
        pAry(LneCnt, 2) = ""
        pAry(LneCnt, 3) = arrRecords2(1, PrtCnt) & ""                  ' 牌号
        pAry(LneCnt, 4) = ""
        pAry(LneCnt, 5) = arrRecords2(2, PrtCnt) & ""                  ' 规格
        pAry(LneCnt, 6) = ""
        pAry(LneCnt, 7) = arrRecords2(3, PrtCnt) & ""                  ' 件数
        pAry(LneCnt, 8) = arrRecords2(4, PrtCnt) & ""                  ' 重量
        pAry(LneCnt, 9) = arrRecords2(5, PrtCnt) & ""                  ' 证明书号
        pAry(LneCnt, 10) = arrRecords2(6, PrtCnt) & ""                 ' 轧制号
       
        If LneCnt = 25 Then
            
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
        
            xlApp.Workbooks.Open (App.Path & "\AHC091C.xls")
            Set xlSheet = xlApp.Worksheets("Sheet1")
            
'            If LneCnt > 1 Then
'                Call MillSheetPrint_C_Line(LneCnt)
'            End If
            
            Call MillSheetPrint_C_Head_1
                        
            sRow = "A" & 18 & ":J" & LneCnt + 17
            xlSheet.Range(sRow).Value = pAry
'            xlSheet.Range("A10:M63").Value = pAry

            xlApp.Range(sRow).Select
            With xlApp.Selection.Borders
                .LineStyle = 1
            End With

            
            xlApp.Application.Visible = True
'            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'            Set xlSheet = Nothing
'            xlApp.ActiveWorkbook.Close False
'            xlApp.Quit
            LneCnt = 0
            
            ReDim pAry(1 To 25, 1 To 10)
            
        End If

    Loop Until PrtCnt = RowCnt
    
    If LneCnt <> 0 Then
    
        
        Set xlApp = GetObject("", "Excel.Application")
        If Err.Number = 429 Then
            Set xlApp = CreateObject("", "Excel.Application")
        End If
    
        xlApp.Workbooks.Open (App.Path & "\AHC091C.xls")
        Set xlSheet = xlApp.Worksheets("Sheet1")
        
        sRow = "A" & 18 & ":J" & LneCnt + 17
        xlSheet.Range(sRow).Value = pAry
        
        xlApp.Range(sRow).Select
        With xlApp.Selection.Borders
            .LineStyle = 1
        End With
        
'        If LneCnt > 1 Then
'            Call MillSheetPrint_C_Line(LneCnt)
'        End If
        
        Call MillSheetPrint_C_Head_1
        xlApp.Application.Visible = True
        
'        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'        Set xlSheet = Nothing
'        xlApp.ActiveWorkbook.Close False
'        xlApp.Quit

    End If
            
    Exit Sub
    
    
End Sub

Private Sub MillSheetPrint_C_Head_1()

    xlSheet.Range("B1").Value = arrRecords1(0, 0) & ""        '品种名称
    xlSheet.Range("B4").Value = arrRecords1(1, 0) & ""        '执行标准
    xlSheet.Range("B7").Value = arrRecords1(2, 0) & ""        '生产许可证编号
    xlSheet.Range("B10").Value = arrRecords1(3, 0) & ""       '收货单位
    xlSheet.Range("H10").Value = arrRecords1(4, 0) & ""       '交货状态
    xlSheet.Range("H1").Value = arrRecords1(5, 0) & ""        '合同编号
    xlSheet.Range("H4").Value = arrRecords1(6, 0) & ""        '提单号
    xlSheet.Range("H7").Value = arrRecords1(7, 0) & ""        '车辆号
    xlSheet.Range("B13").Value = arrRecords1(8, 0) & ""       '目的地
    xlSheet.Range("H13").Value = arrRecords1(9, 0) & ""       '日期
    
    
End Sub

'-----------------------------------非船板-------------------------------------------------
Private Function subGetOracleData() As Boolean
    
'    Dim sQuery As String
'    Dim AdoRs As adodb.Recordset
'    Dim arrRecords1 As Variant      'sQueryHead
'    Dim arrRecords2 As Variant      'sQueryDetail
    
    On Error GoTo Err_Track
                
    Set adoRs = New ADODB.Recordset
    
    sQuery = "{call AHC0051C.P_REFER_HEAD("
    sQuery = sQuery + "'" + LBL_SHP_IST_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_TRNS_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_PROD_CD.Caption + "')}"

'-----------------------------------------------------------------------------
        
    adoRs.Open sQuery, M_CN1, adOpenKeyset

    If adoRs.EOF Then GoTo Err_Track
       
    arrRecords1 = adoRs.GetRows
    adoRs.Close
    
    sQuery = "{call AHC0051C.P_REFER_DETAIL("
    sQuery = sQuery + "'" + LBL_SHP_IST_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_TRNS_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_PROD_CD.Caption + "')}"
                                
    adoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If adoRs.EOF Then GoTo Err_Track
    
    arrRecords2 = adoRs.GetRows
    adoRs.Close
    Set adoRs = Nothing
    
'    Call subMdbUpdate(arrRecords1, arrRecords2)
    Call MillSheetPrint_C
    subGetOracleData = True
    
    Exit Function
    
Err_Track:
        
    If IsObject(adoRs) = True Then
        Set adoRs = Nothing
    End If
        
End Function


Private Sub MillSheetPrint_C()
    Dim RowCnt      As Long
    Dim PrtCnt      As Long
    Dim LneCnt      As Long
    Dim pAry()      As String
    Dim sRow        As String
    Dim sRow2       As String
    
    If IsEmpty(arrRecords1) Or IsEmpty(arrRecords2) Then Exit Sub
    
    RowCnt = UBound(arrRecords2, 2)
    
    PrtCnt = -1
    LneCnt = 0
    
    ReDim pAry(1 To 45, 1 To 14)
    
    Do

        LneCnt = LneCnt + 1
        PrtCnt = PrtCnt + 1
        
        pAry(LneCnt, 1) = PrtCnt + 1                                   ' 序号
        pAry(LneCnt, 2) = arrRecords2(0, PrtCnt) & ""                  ' 产品号
        pAry(LneCnt, 3) = ""
        pAry(LneCnt, 4) = arrRecords2(1, PrtCnt) & ""                  ' 钢种
        pAry(LneCnt, 5) = ""
        pAry(LneCnt, 6) = arrRecords2(2, PrtCnt) & ""                  ' 标准
        pAry(LneCnt, 7) = ""
        pAry(LneCnt, 8) = arrRecords2(3, PrtCnt) & ""                  ' 产品等级
        pAry(LneCnt, 9) = arrRecords2(4, PrtCnt) & ""                  ' 厚度
        pAry(LneCnt, 10) = arrRecords2(5, PrtCnt) & ""                 ' 宽度
        pAry(LneCnt, 11) = arrRecords2(6, PrtCnt) & ""                 ' 长度
        pAry(LneCnt, 12) = arrRecords2(7, PrtCnt) & ""                 ' 重量
        pAry(LneCnt, 13) = arrRecords2(8, PrtCnt) & ""                 ' 数量
        pAry(LneCnt, 14) = arrRecords2(9, PrtCnt) & ""                 ' 轧制号
        
       
        If LneCnt = 45 Then
            
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
        
            xlApp.Workbooks.Open (App.Path & "\AHC090C.xls")
            Set xlSheet = xlApp.Worksheets("Sheet1")
            
'            If LneCnt > 1 Then
'                Call MillSheetPrint_C_Line(LneCnt)
'            End If
            
            Call MillSheetPrint_C_Head
                        
            sRow = "A" & 10 & ":N" & LneCnt + 9
            xlSheet.Range(sRow).Value = pAry
'            xlSheet.Range("A10:M63").Value = pAry

            xlApp.Range(sRow).Select
            With xlApp.Selection.Borders
                .LineStyle = 1
            End With

            
            xlApp.Application.Visible = True
'            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'            Set xlSheet = Nothing
'            xlApp.ActiveWorkbook.Close False
'            xlApp.Quit
            LneCnt = 0
            
            ReDim pAry(1 To 45, 1 To 14)
            
        End If

    Loop Until PrtCnt = RowCnt
    
    If LneCnt <> 0 Then
    
        
        Set xlApp = GetObject("", "Excel.Application")
        If Err.Number = 429 Then
            Set xlApp = CreateObject("", "Excel.Application")
        End If
    
        xlApp.Workbooks.Open (App.Path & "\AHC090C.xls")
        Set xlSheet = xlApp.Worksheets("Sheet1")
        
        sRow = "A" & 10 & ":N" & LneCnt + 9
        xlSheet.Range(sRow).Value = pAry
        
        xlApp.Range(sRow).Select
        With xlApp.Selection.Borders
            .LineStyle = 1
        End With
        
'        If LneCnt > 1 Then
'            Call MillSheetPrint_C_Line(LneCnt)
'        End If
        
        Call MillSheetPrint_C_Head
        xlApp.Application.Visible = True
        
'        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'        Set xlSheet = Nothing
'        xlApp.ActiveWorkbook.Close False
'        xlApp.Quit

    End If
            
    Exit Sub
    
    
End Sub

Private Sub MillSheetPrint_C_Head()

    xlSheet.Range("L5").Value = arrRecords1(0, 0) & ""        '明细号
    xlSheet.Range("C4").Value = arrRecords1(1, 0) & ""        '产品
    xlSheet.Range("H4").Value = arrRecords1(2, 0) & ""        '客户
    xlSheet.Range("L4").Value = arrRecords1(3, 0) & ""        '发放日期
    xlSheet.Range("C5").Value = arrRecords1(4, 0) & ""        '仓库
    xlSheet.Range("F5").Value = arrRecords1(5, 0) & ""        '提单号
    xlSheet.Range("H5").Value = arrRecords1(6, 0) & ""        '订单客户
    xlSheet.Range("C6").Value = arrRecords1(7, 0) & ""        '订单号
    xlSheet.Range("F6").Value = arrRecords1(8, 0) & ""        '订单量
    xlSheet.Range("H6").Value = arrRecords1(9, 0) & ""        '目的地
    xlSheet.Range("L6").Value = arrRecords1(10, 0) & ""       '车辆号
    xlSheet.Range("C7").Value = arrRecords1(11, 0) & ""       '总重量
    xlSheet.Range("F7").Value = arrRecords1(12, 0) & ""       '总数量
    xlSheet.Range("H7").Value = arrRecords1(13, 0) & ""       '运输公司
    xlSheet.Range("L7").Value = arrRecords1(14, 0) & ""       '业务人员
    
    
End Sub
'Private Sub MillSheetPrint_C_Line(LneCnt As Long)
'    Dim iDx         As Integer
'    Dim sRow        As String
''    LneCnt = 20
'    For iDx = 2 To LneCnt
'        xlApp.Rows("10:10").Select
'        xlApp.Selection.Copy
'        xlApp.Selection.Insert Shift:=1
'    Next iDx
'    sRow = 30 & ":" & 10 + LneCnt
'    xlApp.Rows(sRow).Select
'    xlApp.Selection.Delete Shift:=1
'
'End Sub


Private Sub Command1_Click()

    On Error GoTo Err_Track
                
    Set adoRs = New ADODB.Recordset
    
    sQuery = "{call AHG0031C.P_REFER_HEAD("
    sQuery = sQuery + "'" + LBL_SHP_IST_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_TRNS_NO.Caption + "',"
    sQuery = sQuery + "'" + LBL_PROD_CD.Caption + "')}"

'-----------------------------------------------------------------------------
        
    adoRs.Open sQuery, M_CN1, adOpenKeyset

    If adoRs.EOF Then GoTo Err_Track
       
    arrRecords3 = adoRs.GetRows
    adoRs.Close
    
    Set adoRs = Nothing
    
    Set xlApp = GetObject("", "Excel.Application")
    If Err.Number = 429 Then
        Set xlApp = CreateObject("", "Excel.Application")
    End If

    xlApp.Workbooks.Open (App.Path & "\AHG010C.xls")
    Set xlSheet = xlApp.Worksheets("Sheet1")
    
'            If LneCnt > 1 Then
'                Call MillSheetPrint_C_Line(LneCnt)
'            End If
    
'    Call MillSheetPrint_C_Head
                
    
    xlApp.Application.Visible = True
'            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'            Set xlSheet = Nothing
'            xlApp.ActiveWorkbook.Close False
'            xlApp.Quit
'            LneCnt = 0
    xlSheet.Range("C5").Value = arrRecords3(0, 0) & ""        '供货单位
    xlSheet.Range("C7").Value = arrRecords3(1, 0) & ""        '标准
    xlSheet.Range("C8").Value = arrRecords3(2, 0) & ""        '尺寸
    xlSheet.Range("H5").Value = arrRecords3(3, 0) & ""        '收货单位
    xlSheet.Range("H6").Value = arrRecords3(4, 0) & ""        '提单号
    xlSheet.Range("H7").Value = arrRecords3(5, 0) & ""        '车号
    xlSheet.Range("H8").Value = LBL_TRNS_NO.Caption & ""      '明细号
    xlSheet.Range("F3").Value = arrRecords3(6, 0) & ""        '日期
Err_Track:
        
    If IsObject(adoRs) = True Then
        Set adoRs = Nothing
    End If

End Sub

Private Sub Form_Activate()

    If LBL_PROD_CD.Caption = "SL" Then
       Command1.Enabled = True
    Else
       Command1.Enabled = False
    End If


On Error GoTo Err_Track
   Dim sQuery As String
   Dim adoRs As ADODB.Recordset
   
   Set adoRs = New ADODB.Recordset
   
   sQuery = " select shp_ist_no,trns_no from "
   If LBL_PROD_CD.Caption = "SL" Then
       sQuery = sQuery + " fp_slab "
   ElseIf LBL_PROD_CD.Caption = "HC" Then
       sQuery = sQuery + " gp_coil "
   ElseIf LBL_PROD_CD.Caption = "PP" Then
       sQuery = sQuery + " gp_plate "
   End If
   sQuery = sQuery + " where shp_ist_no='" + LBL_SHP_IST_NO.Caption + "' group by shp_ist_no,trns_no"
   
   If Gf_Only_Display(M_CN1, sc1, sQuery) Then
   End If

   
   
Err_Track:
        
    If IsObject(adoRs) = True Then
        Set adoRs = Nothing
    End If
   
End Sub

Private Sub Form_Load()
    
    If LBL_PROD_CD.Caption = "SL" Then
       Command1.Enabled = True
    Else
       Command1.Enabled = False
    End If
    
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    sc1.Add Item:=ss1, Key:="Spread"
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    ss1.Col = 1
    ss1.ColHidden = True
    
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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    If ss1.MaxRows < 1 Then
       Exit Sub
    End If
    ss1.Row = Row
    ss1.Col = 2
    LBL_TRNS_NO.Caption = ss1.Text
    
    If LBL_PROD_CD.Caption = "SL" Then
        Dim adoRs As ADODB.Recordset
        Dim sQuery      As String
        
        cmdprint.Enabled = True
        
        Set adoRs = New ADODB.Recordset
           
        sQuery = "SELECT LOAD_WGT FROM HP_LOAD_WGT "
        sQuery = sQuery & "WHERE SHP_IST_NO = '" & LBL_SHP_IST_NO.Caption & "'"
        sQuery = sQuery & "AND TRNS_NO = '" & LBL_TRNS_NO.Caption & "'"
        
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        If Not adoRs.BOF And Not adoRs.EOF Then
           If Val(adoRs.Fields(0)) = 0 Then
              Call MsgBox("产品还未称重(无计量重量)，不能打印发货明细表", vbExclamation + vbOKOnly, "警告")
              cmdprint.Enabled = False
           End If
        Else:
            Call MsgBox("产品还未称重(无计量重量)，不能打印发货明细表", vbExclamation + vbOKOnly, "警告")
            cmdprint.Enabled = False
        End If
        adoRs.Close
        Set adoRs = Nothing
    End If
        
End Sub
