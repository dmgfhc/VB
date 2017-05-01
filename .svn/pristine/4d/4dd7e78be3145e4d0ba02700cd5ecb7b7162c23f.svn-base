Attribute VB_Name = "basQ"
Option Explicit

Public Const SS_TEXTTIP_FIXEDFOCUSONLY = 3
Public Const SS_TEXTTIP_FLOATINGFOCUSONLY = 4

Global sOrderNo As String                    '订单号 OrderNo
Global sOrderItem As String                  '订单序列号 OrderItem
Global sSampCd As String                     '取样代码
Global sSampSearch As String                 '取样代码 Search Text

Global arrSampCd1() As Variant               '取样代码1
Global arrSampCd2() As Variant               '取样代码2
Global arrSampCd3() As Variant               '取样代码3
Global arrSampCd4() As Variant               '取样代码4
Global arrSampCd5() As Variant               '取样代码5
Global sEXLSavePATH As String                '质保书保存路径
'----------------------------------------------------------------------------------------
' 质量证明书用变量
'----------------------------------------------------------------------------------------
Private xlApp       As Object
Private xlSheet     As Object

'Private arrRecords1 As Variant      'sQueryHead
'Private arrRecords2 As Variant      'sQueryDetail - Chem
'Private arrRecords3 As Variant      'sQueryDetail - Mart


'---------------------------------------------------------------------------------------
'   1.ID           : GF_GetCellMaxLength
'   2.Name         : Get Spread MaxLength
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Get Spread MaxLength
'---------------------------------------------------------------------------------------
Public Function GF_GetCellMaxLength(vSP As vaSpread, iRow As Long, iCol As Long) As Double
    With vSP
        .Row = iRow
        .Col = iCol
        GF_GetCellMaxLength = .TypeNumberMax
    End With
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_GET_CELL_VALUE
'   2.Name         : Get Spread Cell Value
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Get Spread Cell Text
'---------------------------------------------------------------------------------------
Public Function Gf_Get_Cell_Value(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With vSP
        .Row = iRow
        .Col = iCol
        Gf_Get_Cell_Value = .Value
    End With
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_GET_CELL_VALUE2
'   2.Name         : Get Spread Cell Value
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Get Spread Cell Text
'---------------------------------------------------------------------------------------
Public Function GF_GET_CELL_VALUE2(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With vSP
        .Row = iRow
        .Col = iCol
        GF_GET_CELL_VALUE2 = Val(.Value)
        If GF_GET_CELL_VALUE2 = 0 Then GF_GET_CELL_VALUE2 = ""
    End With
End Function



'---------------------------------------------------------------------------------------
'   1.ID           : Gf_GetCellText
'   2.Name         : Get Spread Cell Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Get Spread Cell Text
'---------------------------------------------------------------------------------------
Public Function Gf_GetCellText(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With vSP
        .Row = iRow
        .Col = iCol
        Gf_GetCellText = .Text
        If Gf_GetCellText = "0" Then Gf_GetCellText = ""
    End With
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_GetCellText
'   2.Name         : Get Spread Cell Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Get Spread Cell Text
'---------------------------------------------------------------------------------------
Public Function Gf_GetCellNullCheck(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With vSP
        .Row = iRow
        .Col = iCol
        Gf_GetCellNullCheck = NullCheck(.Text, "")
    End With
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : Gf_GetCellText
'   2.Name         : Get Spread Cell Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Get Spread Cell Text
'---------------------------------------------------------------------------------------
Public Function Gf_GetCellText2(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    Dim str As String
    With vSP
        .Row = iRow
        .Col = iCol
        str = Trim(.Text)
        
        If str = "0" Then str = ""
        If str = "/" Then str = ""
        If str = "//" Then str = ""
        If str = "///" Then str = ""
        If str = "////" Then str = ""
        
        If str <> "" And InStr(str, "/") = Len(str) Then
            str = Left(str, Len(str) - 1)
        End If
        
        Gf_GetCellText2 = str
        
    End With
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_SetRowColor
'   2.Name         : Set Spread Row Backcolor
'   3.Input  Value : Spread Name , Row , Color
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Set Spread Row Backcolor
'---------------------------------------------------------------------------------------
Public Sub Gp_SetRowColor(vSP As vaSpread, ByVal iRow As Long)

    Dim i As Long

    Call Gp_Sp_EvenRowBackcolor(vSP)

    With vSP
        .Row = iRow
        For i = 1 To .MaxCols
            .Col = i
            .BackColor = RGB(180, 200, 240)
        Next i
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_SetCellFormula
'   2.Name         : Set Spread Formula
'   3.Input  Value : Spread Name , Row , Col , Formula
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Set Spread Formula
'---------------------------------------------------------------------------------------
Public Sub Gp_SetCellFormula(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long, ByVal sFormula As String)
    With vSP
        .Row = iRow
        .Col = iCol
        .Formula = sFormula
    End With
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GP_SET_CELL_VALUE
'   2.Name         : Set Spread Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Set Spread Text
'---------------------------------------------------------------------------------------
Public Sub GP_SET_CELL_VALUE(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long, sText As Variant)
    
    If iRow <= 0 Then Exit Sub
    
    With vSP
        .Row = iRow
        .Col = iCol
        .Text = NullCheck(sText, "")
    End With
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GP_SET_CELL_VALUE
'   2.Name         : Set Spread Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Set Spread Text
'---------------------------------------------------------------------------------------
Public Sub GP_SET_CELL_VALUE2(vSP As vaSpread, ByVal iRow As Long, ByVal iCol As Long, ByVal iCnt As Integer, sText1 As Variant, Optional ByVal sText2 As Variant, Optional ByVal sText3 As Variant, Optional ByVal sText4 As Variant, Optional ByVal sText5 As Variant, Optional ByVal sText6 As Variant, Optional ByVal sText7 As Variant, Optional ByVal sText8 As Variant, Optional ByVal sText9 As Variant, Optional ByVal sText10 As Variant)
    
    Dim str As String
    Dim TXT_CNT As Integer
    
    If iRow <= 0 Then Exit Sub
    
    With vSP
        .Row = iRow
        .Col = iCol
        
        If iCnt >= 1 Then If sText1 = "0" Or sText1 = " " Then sText1 = ""
        
        If iCnt >= 2 Then If sText2 = "0" Or sText2 = " " Then sText2 = ""
        
        If iCnt >= 3 Then If sText3 = "0" Or sText3 = " " Then sText3 = ""
        
        If iCnt >= 4 Then If sText4 = "0" Or sText4 = " " Then sText4 = ""
        
        If iCnt >= 5 Then If sText5 = "0" Or sText5 = " " Then sText5 = ""
  
        If iCnt >= 6 Then If sText6 = "0" Or sText6 = " " Then sText6 = ""
        
        If iCnt >= 7 Then If sText7 = "0" Or sText7 = " " Then sText7 = ""
        
        If iCnt >= 8 Then If sText8 = "0" Or sText8 = " " Then sText8 = ""
        
        If iCnt >= 9 Then If sText9 = "0" Or sText9 = " " Then sText9 = ""
        
        If iCnt >= 10 Then If sText10 = "0" Or sText10 = " " Then sText10 = ""
        
        Select Case iCnt
        
            Case 2
                 str = sText1 & "/" & sText2
                 If str = "/" Then str = ""
            Case 3
                 str = sText1 & "/" & sText2 & "/" & sText3
                 If str = "//" Then str = ""
            Case 4
                 str = sText1 & "/" & sText2 & "/" & sText3 & "/" & sText4
                 If str = "///" Then str = ""
            Case 5
                 str = sText1 & "/" & sText2 & "/" & sText3 & "/" & sText4 & "/" & sText5
                 If str = "////" Then str = ""
            Case 6
                 str = sText1 & "/" & sText2 & "/" & sText3 & "/" & sText4 & "/" & sText5 & "/" & sText6
                 If str = "/////" Then str = ""
            Case 7
                 str = sText1 & "/" & sText2 & "/" & sText3 & "/" & sText4 & "/" & sText5 & "/" & sText6 & "/" & sText7
                 If str = "//////" Then str = ""
            Case 8
                 str = sText1 & "/" & sText2 & "/" & sText3 & "/" & sText4 & "/" & sText5 & "/" & sText6 & "/" & sText7 & "/" & sText8
                 If str = "///////" Then str = ""
            Case 9
                 str = sText1 & "/" & sText2 & "/" & sText3 & "/" & sText4 & "/" & sText5 & "/" & sText6 & "/" & sText7 & "/" & sText8 & "/" & sText9
                 If str = "////////" Then str = ""
            Case 10
                 str = sText1 & "/" & sText2 & "/" & sText3 & "/" & sText4 & "/" & sText5 & "/" & sText6 & "/" & sText7 & "/" & sText8 & "/" & sText9 & "/" & sText10
                 If str = "/////////" Then str = ""
                 
            Case Else
                 str = sText1
        End Select
        
        If str <> "" And InStr(str, "/") = Len(str) Then
            str = Left(str, Len(str) - 1)
        End If

        .Text = str
    End With
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GP_ChemCode_RowHeader_Clear
'   2.Name         : Set Spread Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Set Spread Text
'---------------------------------------------------------------------------------------
Public Sub GP_ChemCode_RowHeader_Clear(ByVal sFormName As String, vSP As vaSpread, ByVal iRow As Integer, ByVal iCol As Integer)
    
    Dim i As Long
    Dim iColNo As Integer
    Dim iCol1 As Integer
    Dim iCol2 As Integer
    Dim iCol3 As Integer
    
   Select Case sFormName
        
        Case "AQA0020C"
            iCol1 = 0       'code1
            iCol2 = 9      'code2
            iCol3 = 14      'code3
                                      
        Case "AQA0090C"
            iCol1 = 0
            iCol2 = 6
            iCol3 = 11
            
        Case "AQA0130C"
            iCol1 = 0
            iCol2 = 7
            iCol3 = 13
            
        Case "AQB0110C"
            iCol1 = 0
            iCol2 = 9
            iCol3 = 15
          
        Case Else
        
            Exit Sub
                
        End Select
    
       If iCol3 < iCol Then
            iColNo = iCol3
            
       ElseIf iCol2 < iCol Then
            iColNo = iCol2
       
       ElseIf iCol1 < iCol Then
            iColNo = iCol1
       
       Else
            iColNo = 0
       
       End If
       
       With vSP
        For i = 1 To .MaxRows
                
            .Row = i
            .Col = iCol1
            If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
                .Text = ""
            End If
        Next i
        
        For i = 1 To .MaxRows
                
            .Row = i
            .Col = iCol2
            If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
                .Text = ""
            End If
        Next i
        
        For i = 1 To .MaxRows
                
            .Row = i
            .Col = iCol3
            If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
                .Text = ""
            End If
        Next i
        
        .Row = iRow: .Col = iColNo
        
        If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
            .Text = "▼"
            '.Text = ">>"
        End If
        
         
       End With
       
       'vSP.SetFocus
         
'       Call GP_SetRowHeaderClear(vSP, iRow, iColNo)
       
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GP_SetRowHeaderClear
'   2.Name         : Set Spread Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Set Spread Text
'---------------------------------------------------------------------------------------
Public Sub GP_SetRowHeaderClear(vSP As vaSpread, ByVal iRow As Long, Optional ByVal iCol As Integer = 0)
    
    Dim i As Long
    
    With vSP
    
        For i = 1 To .MaxRows
                
            .Row = i
            .Col = iCol
            If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
                .Text = ""
            End If
        Next i
        
        .Row = iRow: .Col = iCol
        
        If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
            .Text = "▼"  '▽▼◎→◆◇
        End If
    
    End With
End Sub



'---------------------------------------------------------------------------------------
'   1.ID           : GF_NullChange
'   2.Name         : Oracle Null Data -> ""
'   3.Input  Value : Oracle Data
'   4.Return Value : Null Change Value
'   5.Writer       :
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : Oracle Null Data -> ""
'---------------------------------------------------------------------------------------
Public Function GF_NullChange(ByVal strVal As Variant) As Variant

    If IsNull(strVal) = True Then
        GF_NullChange = ""
    Else
        If strVal = "/" Then
            GF_NullChange = strVal
        Else
            GF_NullChange = strVal
        End If
        
    End If
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GP_BACKCOLOR_WHITE
'   2.Name         : Input Control Backcolor = White
'   3.Input  Value : rControl Collection
'   4.Return Value :
'   5.Writer       : Chu Kyo Su
'   6.Create Date  : 2003. 09 .26
'   7.Modify Date  :
'   8.Comment      : Input Control Backcolor = White
'---------------------------------------------------------------------------------------
Public Sub GP_BACKCOLOR_WHITE(iControl As Collection)
    
    Dim icount As Integer

    For icount = 1 To iControl.COUNT
        
        iControl.Item(icount).BackColor = vbWhite
        
    Next icount
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GP_SPREAD_UNLOCK
'   2.Name         : Input Control Backcolor = White
'   3.Input  Value : rControl Collection
'   4.Return Value :
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 09 .26
'   7.Modify Date  :
'   8.Comment      : Input Control Backcolor = White
'---------------------------------------------------------------------------------------
Public Sub GP_SPREAD_UNLOCK(ByVal sp As vaSpread, ByVal iRow As Long)
    With sp
    
        .Col = 0: .Col2 = .MaxCols
        .Row = iRow: .Row2 = -1
        
        .BlockMode = True
        .Lock = False
        .BlockMode = False
       
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GP_SPREAD_LOCK
'   2.Name         : Input Control Backcolor = White
'   3.Input  Value : rControl Collection
'   4.Return Value :
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 09 .26
'   7.Modify Date  :
'   8.Comment      : Input Control Backcolor = White
'---------------------------------------------------------------------------------------
Public Sub GP_SPREAD_LOCK(ByVal sp As vaSpread, ByVal iRow As Long)
    With sp
    
        .Col = 0: .Col2 = .MaxCols
        .Row = iRow: .Row2 = -1
        
        .BlockMode = True
        .Lock = True
        .BlockMode = False
       
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GP_MENU_SHOW_HIDE
'   2.Name         : MenuBar Show and Hide
'   3.Input  Value : rControl Collection
'   4.Return Value :
'   5.Writer       : Chu Kyo Su
'   6.Create Date  : 2003. 09 .26
'   7.Modify Date  :
'   8.Comment      : MenuBar Show and Hide
'---------------------------------------------------------------------------------------
Public Sub GP_MENU_SHOW_HIDE(ByVal sText As String)
    
    Dim i As Integer
    Dim iCnt As Integer
    
    Dim iCol() As Integer
    Dim bChk() As Boolean
    
    iCnt = Len(sText) / 3
    
    ReDim iCol(iCnt)
    ReDim bChk(iCnt)
    
    For i = 1 To iCnt
    
        iCol(i) = CInt(Mid(sText, (i * 3 - 2), 2))
        
        If Mid(sText, (i * 3), 1) = "T" Then
            bChk(i) = True
        Else
            bChk(i) = False
        End If
        
    Next i
    
    For i = 1 To iCnt
        
        MDIMain.MenuTool.Buttons(iCol(i)).Enabled = bChk(i)
        
    Next i
    
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_MS_CommonNameFind
'   2.Name         : Common Code Name Find
'   3.Input  Value : Common Code , Code(TextBox) , CodeName(TextBox),CodeName1(TextBox)
'   4.Return Value :
'   5.Writer       : Chu Kyo Su
'   6.Create Date  : 2003. 10. 10
'   7.Modify Date  :
'   8.Comment      : Matser Type Common Code Name Find
'---------------------------------------------------------------------------------------
Public Sub Gp_MS_CodeNameFind(KeyCode As Integer, ByVal sCode As String, oCode As Object, Optional oCodeName As Object, Optional oCodeName1 As Object)
              
    Dim bType As Boolean
    Dim bCheck As Boolean
              
    If KeyCode = vbKeyF4 Then
                 
        DD.sWitch = "MS"
        DD.rControl.Add Item:=oCode
        DD.nameType = "2"
        
        If Not oCodeName Is Nothing Then DD.rControl.Add Item:=oCodeName
        If Not oCodeName1 Is Nothing Then DD.rControl.Add Item:=oCodeName1
        
        Select Case sCode
        
            Case "STDSPEC"              '标准号
                Call Gf_StdSPEC_DD(M_CN1, KeyCode)
                
            Case "STDSPEC2"             '在用标准
                Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
            
            Case "CUST_CD"              '客户
                DD.nameType = "1"
                Call Gf_Customer_DD(M_CN1, KeyCode)
            
            Case "ENDUSE_CD"            '订单用途
                Call Gf_Usage_DD(M_CN1, KeyCode)
            
            Case "STLGRD"               '钢种
                Call Gf_Stlgrd_DD(M_CN1, KeyCode)
            
            Case "CUST_SPEC_NO"         '客户特殊要求编号
                Call Gf_Cust_STD_DD(M_CN1, KeyCode)
            
            Case "NISCO_QUALITY_NO"     '企标材质编号
                Call Gf_Nisco_STD_DD(M_CN1, KeyCode)
                
            Case "MLT_STD_NO"           '炼钢规程编号
                Call Gf_Melt_STD_DD(M_CN1, KeyCode)
                
            Case "MILL_STD_NO"          '轧钢规程编号
                Call Gf_MILL_STD_DD(M_CN1, KeyCode)
            
            Case "DEV_STD_CD"          '代表性交付条件标准
                Call Gf_STD_DELV_DD(M_CN1, KeyCode)
                
            Case "HTM_COND_CD"
                Call Gf_HEAT_COND_DD(M_CN1, KeyCode)
                            
            Case Else                   'Common Code
                DD.sKey = sCode
                Call Gf_Common_DD(M_CN1, vbKeyF4)
                bCheck = True
        
        End Select
        
    Else    'Max Length Input -> Code Name Find

        'If sCode = "" Then Exit Sub
        If KeyCode = 13 Or KeyCode = 20 Then Exit Sub
        
        If oCodeName Is Nothing Then Exit Sub
        
        Select Case sCode
        
            Case "STDSPEC"      '标准号
                bType = False
            Case Else           'Common Code
                bType = True
        End Select

        If bType = True And Len(Trim(oCode.Text)) = oCode.MaxLength Then
            If Left(oCodeName.Name, 3) = "lbl" And bCheck = True Then
                oCodeName.Caption = ""
                oCodeName.Caption = Gf_ComnNameFind(M_CN1, sCode, oCode.Text, "2")
            ElseIf bType = True Then
                oCodeName.Text = ""
                oCodeName.Text = Gf_ComnNameFind(M_CN1, sCode, oCode.Text, "2")
            End If
        ElseIf Len(Trim(oCode.Text)) = 0 Then
            If Left(oCodeName.Name, 3) = "lbl" Then
                oCodeName.Caption = ""
            Else
                oCodeName.Text = ""
            End If
        End If
        
    End If

End Sub


''---------------------------------------------------------------------------------------
''   1.ID           : Gp_CodeNameFind
''   2.Name         : Code Name Find
''   3.Input  Value : Common Code , Code(TextBox) , CodeName(TextBox)
''   4.Return Value :
''   5.Writer       : Chu Kyo Su
''   6.Create Date  : 2003. 10. 10
''   7.Modify Date  :
''   8.Comment      : Code Name Find
''---------------------------------------------------------------------------------------
'Public Sub Gp_CodeNameFind(ByVal sCode As String, oCode As TextBox, oCodeName As TextBox)
'
'        oCodeName.Text = ""
'        oCodeName.Text = Gf_ComnNameFind(M_CN1, sCode, oCode.Text, "2")
'
'End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Procedure_Exec
'   2.Name         : Oracle Procedure Execute
'   3.Input  Value : Procedure Name
'   4.Return Value : Boolean
'   5.Writer       : Chu Kyo Su
'   6.Create Date  : 2003. 10. 10
'   7.Modify Date  :
'   8.Comment      : Oracle Procedure Execute
'---------------------------------------------------------------------------------------
Public Function Gf_Procedure_Exec(ByVal sQuery As String) As Boolean

On Error GoTo ExecQuery_ERROR

    Dim ret_Result_ErrCode As String
    Dim ret_Result_ErrMsg As String
    Dim adoCmd As adodb.Command
    Dim OutParam(2, 4) As Variant
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    'Db Connection Check
    If M_CN1 Is Nothing Then
        If GF_DbConnect = False Then Gf_Procedure_Exec = False: Exit Function
    End If
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdStoredProc
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If Trim(adoCmd("arg_e_code")) = "Y" Then
        ret_Result_ErrCode = Trim(adoCmd("arg_e_code"))
        ret_Result_ErrMsg = NullCheck(adoCmd("arg_e_msg"), "")
        
        sErrMessg = "错误 代码 : " & ret_Result_ErrCode & vbCrLf & "错误 信息 : " & ret_Result_ErrMsg
        
        Call Gp_MsgBoxDisplay(sErrMessg)
        
        Set adoCmd = Nothing
        Gf_Procedure_Exec = False
    
        Exit Function
    Else
        ret_Result_ErrMsg = NullCheck(adoCmd("arg_e_msg"), "")
        sErrMessg = ret_Result_ErrMsg
        
        Call Gp_MsgBoxDisplay(sErrMessg)
        
    End If
    
    Set adoCmd = Nothing
    Gf_Procedure_Exec = True
    
    Exit Function

ExecQuery_ERROR:

    Set adoCmd = Nothing
    Gf_Procedure_Exec = False
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_QualityCode
'   2.Name         : Quality Common Code Find
'   3.Input  Value : KeyCode,vName, vTag
'   4.Return Value : None
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Quality Common Code Find
'---------------------------------------------------------------------------------------
Public Sub GF_QualityCode(KeyCode As Integer, ByVal vName As String, ByVal vTag As String, ByVal oForm As Form)
    
    Dim vCDNAME As String
    Dim vKey As String
    Dim vLen As String

    If vTag = "N" Or vTag = "EMP_ID" Or vTag = "" Then
        Exit Sub
    End If

    If TypeName(oForm.Controls(vName)) = "TextBox" Then
        If (InStr(vName, "DSC_CD") > 0) Or (InStr(vName, "UNIT") > 0) Then
            vCDNAME = "N"
            vKey = vTag
        Else
            If (InStr(vName, "CD") > 0) Or (InStr(vName, "LOC") > 0) Or (InStr(vName, "TYP") > 0) Or (InStr(vName, "KND") > 0) Then
                'vCDNAME = vName + "NAME"
                vKey = Right(vTag, 5)
                vLen = Len(Trim(vTag)) - 6
                vCDNAME = Left(vTag, vLen)
            End If
        End If

        If KeyCode = vbKeyF4 Then
            If (InStr(vName, "DSC_CD") > 0) Or (InStr(vName, "UNIT") > 0) Then
                DD.sWitch = "MS"
                DD.sKey = vKey
                DD.rControl.Add Item:=oForm.Controls(vName)
                DD.nameType = "1"
                Call Gf_Common_DD(M_CN1, KeyCode)
                Exit Sub
            Else
                If (InStr(vName, "CD") > 0) Or (InStr(vName, "LOC") > 0) Or (InStr(vName, "TYP") > 0) Or (InStr(vName, "KND") > 0) Then
                   '
                    DD.sWitch = "MS"
                    DD.sKey = vKey
                    DD.rControl.Add Item:=oForm.Controls(vName)
                    DD.rControl.Add Item:=oForm.Controls(vCDNAME)
                    DD.nameType = "2"
                    Call Gf_Common_DD(M_CN1, KeyCode)
                    Exit Sub
                End If
            End If
        End If

        If vCDNAME <> "N" Then
            If Len(Trim(oForm.Controls(vName).Text)) = oForm.Controls(vName).MaxLength Then
                oForm.Controls(vCDNAME).Text = Gf_ComnNameFind(M_CN1, vKey, Trim(oForm.Controls(vName).Text), 2)
            Else
                oForm.Controls(vCDNAME).Text = ""
            End If
        End If
    End If
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GP_QUALITY_CODE
'   2.Name         : Quality Common Code Find
'   3.Input  Value : KeyCode,vName, vTag , Form , iType
'   4.Return Value : None
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : Quality Common Code Find
'---------------------------------------------------------------------------------------
Public Sub GP_QUALITY_CODE(KeyCode As Integer, ByVal vName As String, ByVal vTag As String, ByVal oForm As Form)
    
    Dim vCDNAME As String
    Dim vKey As String
    Dim vLen As String

    If vTag = "" Then Exit Sub
       
    'vCDNAME = vName + "NAME"
    vKey = Right(vTag, 5)
    vLen = Len(Trim(vTag)) - 6
    If vLen < 0 Then
        vCDNAME = "N"
    Else
        vCDNAME = Left(vTag, vLen)
    End If
    
    DD.sWitch = "MS"
    DD.sKey = vKey
    DD.rControl.Add Item:=oForm.Controls(vName)
    DD.nameType = "2"
    Call Gf_Common_DD(M_CN1, KeyCode)

    If vCDNAME <> "N" Then

        If Len(Trim(oForm.Controls(vName).Text)) = oForm.Controls(vName).MaxLength Then
            oForm.Controls(vCDNAME).Text = Gf_ComnNameFind(M_CN1, vKey, Trim(oForm.Controls(vName).Text), 2)
        Else
            oForm.Controls(vCDNAME).Text = ""
        End If
    
    End If
'
End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : GS_Combo_THK_MAX
'   2.Name         : THK_MAX ADD TO COMBOBOX
'   3.Input  Value :
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : THK_MAX ADD TO COMBOBOX
'---------------------------------------------------------------------------------------
Public Sub GS_Combo_THK_MAX(ByVal oForm As Form)

On Error GoTo Error_Rtn

    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i As Integer
    Dim sTable As String
    
    If Trim(oForm.txt_STDSPEC.Text) = "" Or Trim(oForm.txt_STDSPEC_YY.Text) = "" Then
        Exit Sub
    End If

    oForm.cbo_THK_MAX.Clear
    oForm.cbo_THK_MIN.Clear
    
    Screen.MousePointer = vbHourglass
    
    Set AdoRs = New adodb.Recordset
    
    Select Case oForm.Name
    
        Case "AQA0020C"
            sTable = "QP_STD_CHEM"
        
        Case "AQA0030C"
            sTable = "QP_STD_MATR"
        
        Case "AQA0140C"
            sTable = "QP_NISCO_MATR"
    
        Case Else
            sTable = "QP_STD_CHEM"
    End Select
   
    sQuery = "SELECT  DISTINCT THK_MIN , THK_MAX FROM " + sTable + " WHERE STDSPEC = '" + Trim(oForm.txt_STDSPEC.Text) + "' AND STDSPEC_YY = '"
    sQuery = sQuery + Trim(oForm.txt_STDSPEC_YY.Text) + "'"
 
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.EOF Then

'        oForm.cbo_THK_MAX.Clear
'        oForm.cbo_THK_MIN.Clear
        
        ArrayRecords = AdoRs.GetRows
        
        For i = 0 To UBound(ArrayRecords, 2)
            
            oForm.cbo_THK_MIN.AddItem ArrayRecords(0, i)
            oForm.cbo_THK_MAX.AddItem ArrayRecords(1, i)
        Next i
                                    
    End If
                
    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    
Error_Rtn:
    Screen.MousePointer = vbDefault

End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GS_Combo_THK_MAX
'   2.Name         : THK_MAX ADD TO COMBOBOX
'   3.Input  Value :
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : THK_MAX ADD TO COMBOBOX
'---------------------------------------------------------------------------------------
Public Sub GS_Combo_THK_MAX2(ByVal oForm As Form)

On Error GoTo Error_Rtn

    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i As Integer
    Dim sTable As String
    Dim iHeight As Integer
    
    If Trim(oForm.txt_STDSPEC.Text) = "" Or Trim(oForm.txt_STDSPEC_YY.Text) = "" Then
        Exit Sub
    End If
    
    iHeight = 245
    
    Screen.MousePointer = vbHourglass
    
    Set AdoRs = New adodb.Recordset
    
    Select Case oForm.Name
    
        Case "AQA0020C"
            sTable = "QP_STD_CHEM"
        
        Case "AQA0030C"
            sTable = "QP_STD_MATR"
        
        Case "AQA0140C"
            sTable = "QP_NISCO_MATR"
    
        Case Else
            sTable = "QP_STD_CHEM"
    End Select
   
    sQuery = "SELECT  DISTINCT THK_MIN , THK_MAX FROM " + sTable + " WHERE STDSPEC = '" + Trim(oForm.txt_STDSPEC.Text) + "' AND STDSPEC_YY = '"
    sQuery = sQuery + Trim(oForm.txt_STDSPEC_YY.Text) + "'"
 
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
     With oForm.ss2
     
        .MaxRows = 0
    
    If Not AdoRs.EOF Then

'        oForm.cbo_THK_MAX.Clear
'        oForm.cbo_THK_MIN.Clear
        
        ArrayRecords = AdoRs.GetRows
                           
        .MaxRows = UBound(ArrayRecords, 2) + 1
     '   .RowHeight = 300
        .Height = (.MaxRows * iHeight) + 15
        
        For i = 0 To UBound(ArrayRecords, 2)
        
            .Row = i + 1
                        
            .Col = 1: .Text = Val(ArrayRecords(0, i))
            .Col = 2: .Text = Val(ArrayRecords(1, i))
        Next i
        
        
    Else
     
     .MaxRows = 1
                                    
    End If
    
'     .MaxRows = 5
'     .Height = 1500
'     .Row = 1
'     .Col = 1
'     .Text = "11"
'     .Col = 2
'     .Text = "12"
    
    End With
                
    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    
Error_Rtn:
    Screen.MousePointer = vbDefault

End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Thick_Mix_Max
'   2.Name         : THK_MAX , THK_MIN ADD TO COMBOBOX
'   3.Input  Value :
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : THK_MAX ADD TO COMBOBOX
'---------------------------------------------------------------------------------------
Public Function Gf_Thick_Mix_Max(ByVal oForm As Form) As Variant

On Error GoTo Error_Rtn

    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i As Integer
    Dim sTable As String
        
    Screen.MousePointer = vbHourglass
    
    Set AdoRs = New adodb.Recordset
    
    Select Case oForm.Name
    
        Case "AQA0020C"
            sTable = "QP_STD_CHEM"
        
        Case "AQA0030C"
            sTable = "QP_STD_MATR"
        
        Case "AQA0140C"
            sTable = "QP_NISCO_MATR"
    
        Case Else
            sTable = "QP_STD_CHEM"
    End Select
   
    sQuery = "SELECT  DISTINCT THK_MIN , THK_MAX FROM " + sTable + " WHERE STDSPEC = '" + Trim(oForm.txt_STDSPEC.Text) + "' AND STDSPEC_YY = '"
    sQuery = sQuery + Trim(oForm.txt_STDSPEC_YY.Text) + "'"
 
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.EOF Then

        ArrayRecords = AdoRs.GetRows
                                                                   
    End If
    
    Gf_Thick_Mix_Max = ArrayRecords
    
'    End With
                
    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    
Error_Rtn:
    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GP_SET_THK_MIN_MAX_VALUE
'   2.Name         : 厚度组 MIN , MAX VALUE
'   3.Input  Value : MIN , MAX , TAEGET
'   4.Return Value : None
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .30
'   7.Modify Date  :
'   8.Comment      : 厚度组 MIN , MAX VALUE
'---------------------------------------------------------------------------------------
Public Sub GP_SET_THK_MIN_MAX_VALUE(ByVal oForm As Form)

    Dim iSeq As Integer
    Dim sText As String
    Dim sMin As Double
    Dim sMax As Double
    
    oForm.cbo_THK_MIN_MAX.Text = Trim(oForm.cbo_THK_MIN_MAX.Text)
    
    sText = Trim(oForm.cbo_THK_MIN_MAX.Text)
    
    sText = Replace(sText, "Min:", "")
    sText = Replace(sText, "Max:", "")
    
    iSeq = InStr(1, Trim(sText), "/")
    
    If sText = "" Then
        sMin = 0: sMax = 0
    Else
        sMin = Val(Left(sText, iSeq - 1))
        sMax = Val(Mid(sText, iSeq + 1))
    End If
    oForm.txt_THK_MIN.Value = sMin
    oForm.txt_THK_MAX.Value = sMax

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GF_MIN_MAX_TARGET_CHECK
'   2.Name         : 下限值 , 上限值 , 目标值 Check
'   3.Input  Value : MIN , MAX , TAEGET
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : 下限值 , 上限值 , 目标值 Check
'---------------------------------------------------------------------------------------
Public Function GF_MIN_MAX_TARGET_CHECK(ByVal dMin As Object, ByVal dMax As Object, ByVal dTgt As Object) As Boolean

    If Trim(dMin.Text) = "" And Trim(dMax.Text) = "" And Trim(dTgt.Text) = "" Then
        GF_MIN_MAX_TARGET_CHECK = True
        Exit Function
    End If

    If Trim(dMin.Text) <= Trim(dTgt.Text) And Trim(dMax.Text) >= Trim(dTgt.Text) Then
        GF_MIN_MAX_TARGET_CHECK = True
    Else
        GF_MIN_MAX_TARGET_CHECK = False
        Call Gp_MsgBoxDisplay("输入数据错误,超出范围!", "I")
        dMin.SetFocus
    End If
                
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : GF_CHEM_SEQ
'   2.Name         : CHEMICAL Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : CHEMICAL Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function GF_CHEM_SEQ(Conn As adodb.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    Dim i As Long
    Dim dblChem As Double
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.COUNT = 0 Or DD.rControl.COUNT > 2 Then
        Call Gp_MsgBoxDisplay("数据字典条件无效.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "CHEM"        'Order Usage Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
        
    If DD.sWitch = "SP" Then
    
        DD.sQuery = "SELECT CHEM_COMP_CD , CHEM_COMP_SEQ  , CHEM_LEN FROM QP_CHEM_SEQ "
        
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sWhere = "WHERE CHEM_COMP_CD LIKE '" & Trim(DD.sPname.Text) & "%' ORDER BY CHEM_COMP_SEQ"
                
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
            If DD.rControl.COUNT > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                sNew_Name = DD.sPname.Text
            End If
            
            DD.sPname.TabStop = True
            DD.sPname.SetFocus
            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
            DD.sPname.EditMode = True
            DD.sPname.TabStop = False
            
            If DD.sSelect Then
                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
            End If
        End If
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_GetChemicalCode
'   2.Name         : Get CHEMICAL Code
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : CHEMICAL Code
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : Get CHEMICAL Code
'---------------------------------------------------------------------------------------
Public Function GF_GetChemicalCode() As Variant

On Error GoTo Error_Rtn

    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
        
    Screen.MousePointer = vbHourglass
                
    sQuery = "SELECT CHEM_COMP_SEQ,CHEM_COMP_CD,CHEM_LEN FROM QP_CHEM_SEQ ORDER BY CHEM_COMP_SEQ"
    
    Set AdoRs = New adodb.Recordset
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.BOF Or AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        Screen.MousePointer = 0
        GF_GetChemicalCode = Null
        Exit Function
    End If
        
    GF_GetChemicalCode = AdoRs.GetRows
        
    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

Error_Rtn:

    Set AdoRs = Nothing
    GF_GetChemicalCode = Null

    Screen.MousePointer = vbDefault
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_GetSampleCode
'   2.Name         : 取样代码检索
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : 取样代码检索
'---------------------------------------------------------------------------------------
Public Sub Gp_GetSampleCode()

'On Error GoTo Error_Rtn

    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    Screen.MousePointer = vbHourglass
    
    ReDim arrSampCd1(9, 2)
    
    arrSampCd1(0, 0) = "01"
    arrSampCd1(1, 0) = "02"
    arrSampCd1(2, 0) = "03"
    arrSampCd1(3, 0) = "04"
    arrSampCd1(4, 0) = "05"
    arrSampCd1(5, 0) = "06"
    arrSampCd1(6, 0) = "07"
    arrSampCd1(7, 0) = "08"
    arrSampCd1(8, 0) = "09"
            
    arrSampCd1(0, 1) = "取样1块"
    arrSampCd1(1, 1) = "取样2块"
    arrSampCd1(2, 1) = "取样3块"
    arrSampCd1(3, 1) = "取样4块"
    arrSampCd1(4, 1) = "取样5块"
    arrSampCd1(5, 1) = "取样6块"
    arrSampCd1(6, 1) = "取样7块"
    arrSampCd1(7, 1) = "取样8块"
    arrSampCd1(8, 1) = "取样9块"
                
    Set AdoRs = New adodb.Recordset
    
'长度方向部位
    sQuery = "SELECT CD , CD_NAME FROM ZP_CD WHERE CD_MANA_NO  = 'Q0021'"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF Then
        arrSampCd2 = AdoRs.GetRows
        AdoRs.Close
    End If
        
'宽度方向部位
    sQuery = "SELECT CD , CD_NAME FROM ZP_CD WHERE CD_MANA_NO  = 'Q0022'"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF Then
        arrSampCd3 = AdoRs.GetRows
        AdoRs.Close
    End If
        

'厚度方向部位
    sQuery = "SELECT CD , CD_NAME FROM ZP_CD WHERE CD_MANA_NO  = 'Q0023'"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF Then
        arrSampCd4 = AdoRs.GetRows
        AdoRs.Close
    End If


'试样尺寸代码
    sQuery = "SELECT SMP_SIZE_CD , SMP_SPEC FROM QP_SAMP_STD ORDER BY SMP_SIZE_CD"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF Then
        arrSampCd5 = AdoRs.GetRows
        'AdoRs.Close
    End If
        
    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Error_Rtn:

    Set AdoRs = Nothing

    Screen.MousePointer = vbDefault
    
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : Gp_SetSampleCode
'   2.Name         : 取样代码 Setting
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : 取样代码 Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_SetSampleCode(oForm As Form)

'On Error GoTo Error_Rtn

    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    Dim i As Integer
    
    With oForm
        
        .cbo_Name1.Clear
        .cbo_Name2.Clear
        .cbo_Name3.Clear
        .cbo_Name4.Clear
        .cbo_Name5.Clear
        .cbo_Cd1.Clear
        .cbo_Cd2.Clear
        .cbo_Cd3.Clear
        .cbo_Cd4.Clear
        .cbo_Cd5.Clear
        
        For i = 0 To UBound(arrSampCd1) - 1
        
            .cbo_Cd1.AddItem arrSampCd1(i, 0)
            .cbo_Name1.AddItem arrSampCd1(i, 0) + " : " + arrSampCd1(i, 1)
        
        Next i
    
        For i = 0 To UBound(arrSampCd2, 2)
        
            .cbo_Cd2.AddItem arrSampCd2(0, i)
            .cbo_Name2.AddItem arrSampCd2(0, i) + " : " + arrSampCd2(1, i)
        
        Next i
        
        For i = 0 To UBound(arrSampCd3, 2)
        
            .cbo_Cd3.AddItem arrSampCd3(0, i)
            .cbo_Name3.AddItem arrSampCd3(0, i) + " : " + arrSampCd3(1, i)
        
        Next i
        
        For i = 0 To UBound(arrSampCd4, 2)
        
            .cbo_Cd4.AddItem arrSampCd4(0, i)
            .cbo_Name4.AddItem arrSampCd4(0, i) + " : " + arrSampCd4(1, i)
        
        Next i
        
        For i = 0 To UBound(arrSampCd5, 2)
        
            .cbo_Cd5.AddItem arrSampCd5(0, i)
            If IsNull(arrSampCd5(1, i)) Then
                .cbo_Name5.AddItem arrSampCd5(0, i) + " : " + ""
            Else
                .cbo_Name5.AddItem arrSampCd5(0, i) + " : " + arrSampCd5(1, i)
            End If
        
        Next i
        
    
    End With
       
    Exit Sub

Error_Rtn:

    
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : Gp_SetComboBoxListIndex
'   2.Name         : ComboBox Listindex Setting
'   3.Input  Value : ComboBox , String
'   4.Return Value : ComboBox Listindex Setting
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 10 .07
'   7.Modify Date  : 2003. 10 .07
'   8.Comment      : Get Ceq Value
'---------------------------------------------------------------------------------------
Public Sub Gp_SetComboBoxListIndex(ByVal oCombo As ComboBox, ByVal sSearch As String)

    Dim i As Integer

    With oCombo
    
        For i = 1 To .ListCount
        
            .ListIndex = i - 1
            
            If sSearch = .Text Then
            
                '.ListIndex = i
            
                Exit Sub
            End If
                            
        Next i
        
        .ListIndex = -1
    
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GF_GetCeqValue
'   2.Name         : Get Ceq Value
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : CHEMICAL Code
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : Get Ceq Value
'---------------------------------------------------------------------------------------
Public Sub GF_GetCeqValue(ByVal vSP As vaSpread, ByVal sFormName As String, ByVal sKnd As String)
    
    Dim i As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim arrCol(12) As Integer
    
    Dim iCode1 As Integer
    Dim iCode2 As Integer
    Dim iCode3 As Integer
    
    Dim iMin1 As Integer
    Dim iMin2 As Integer
    Dim iMin3 As Integer
    
    Dim iMax1 As Integer
    Dim iMax2 As Integer
    Dim iMax3 As Integer
    
    Dim iTgt1 As Integer
    Dim iTgt2 As Integer
    Dim iTgt3 As Integer
    
    Dim iSeq1 As Integer
    Dim iSeq2 As Integer
    Dim iSeq3 As Integer
    
    
    Dim FOMULA_CD As String
    
    Dim CHEM_C(3) As Double
    Dim CHEM_SI(3) As Double
    Dim CHEM_MN(3) As Double
    Dim CHEM_P(3) As Double
    Dim CHEM_S(3) As Double
    Dim CHEM_CR(3) As Double
    Dim CHEM_V(3) As Double
    Dim CHEM_MO(3) As Double
    Dim CHEM_CU(3) As Double
    Dim CHEM_NI(3) As Double
    Dim CHEM_B(3) As Double
    
    Dim Chem_Val(3) As Double
    Dim chem_max As Double
    Dim chem_tgt As Double
    
    With vSP
    
        iRow = .Row
        iCol = .ActiveCol
        .Col = iCol
        FOMULA_CD = .Text
        
        Select Case sFormName
        
            Case "AQA0020C"
                iCode1 = 5       'code1
                iCode2 = 10      'code2
                iCode3 = 15      'code3
                iMin1 = 6       'min1
                iMin2 = 11      'min2
                iMin3 = 16      'min3
                iMax1 = 7       'max1
                iMax2 = 12      'max2
                iMax3 = 17      'max3
                
                iSeq1 = iCol - 2
                iSeq2 = iCol - 1
    
                                
            Case "AQA0090C"
                iCode1 = 2
                iCode2 = 7
                iCode3 = 12
                iMin1 = 3
                iMin2 = 8
                iMin3 = 13
                iMax1 = 4
                iMax2 = 9
                iMax3 = 14
                
                iSeq1 = iCol - 2
                iSeq2 = iCol - 1
                
                
                
            Case "AQA0130C"
                iCode1 = 2
                iCode2 = 8
                iCode3 = 14
                iMin1 = 3
                iMin2 = 9
                iMin3 = 15
                iMax1 = 4
                iMax2 = 10
                iMax3 = 16
                iTgt1 = 5
                iTgt2 = 11
                iTgt3 = 17
                
                iSeq1 = iCol - 3
                iSeq2 = iCol - 2
                iSeq3 = iCol - 1
                
            Case "AQB0110C"
                iCode1 = 4
                iCode2 = 10
                iCode3 = 16
                iMin1 = 5
                iMin2 = 11
                iMin3 = 17
                iMax1 = 6
                iMax2 = 12
                iMax3 = 18
                iTgt1 = 7
                iTgt2 = 13
                iTgt3 = 19
                
                iSeq1 = iCol - 3
                iSeq2 = iCol - 2
                iSeq3 = iCol - 1
                 
            Case "AQC0080C"
                iCode1 = 3
                iMin1 = 4
                iMax1 = 5
                
                iSeq1 = 4
                iSeq2 = 5
            Case Else
            
                Exit Sub
                
        End Select
            
        
'-------------------- Min value -----------------------------------------------------------
            
            CHEM_C(0) = subGetChemValue(vSP, "C", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_SI(0) = subGetChemValue(vSP, "Si", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_MN(0) = subGetChemValue(vSP, "Mn", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_P(0) = subGetChemValue(vSP, "P", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_S(0) = subGetChemValue(vSP, "S", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_CR(0) = subGetChemValue(vSP, "Cr", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_V(0) = subGetChemValue(vSP, "V", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_MO(0) = subGetChemValue(vSP, "Mo", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_CU(0) = subGetChemValue(vSP, "Cu", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_NI(0) = subGetChemValue(vSP, "Ni", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_B(0) = subGetChemValue(vSP, "B", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            
'-------------------- Max value -----------------------------------------------------------
            
            CHEM_C(1) = subGetChemValue(vSP, "C", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_SI(1) = subGetChemValue(vSP, "Si", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_MN(1) = subGetChemValue(vSP, "Mn", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_P(1) = subGetChemValue(vSP, "P", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_S(1) = subGetChemValue(vSP, "S", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_CR(1) = subGetChemValue(vSP, "Cr", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_V(1) = subGetChemValue(vSP, "V", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_MO(1) = subGetChemValue(vSP, "Mo", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_CU(1) = subGetChemValue(vSP, "Cu", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_NI(1) = subGetChemValue(vSP, "Ni", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_B(1) = subGetChemValue(vSP, "B", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)

'-------------------- Target value -----------------------------------------------------------
            
        If sKnd = "3" Then
        
            CHEM_C(2) = subGetChemValue(vSP, "C", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_SI(2) = subGetChemValue(vSP, "Si", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_MN(2) = subGetChemValue(vSP, "Mn", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_P(2) = subGetChemValue(vSP, "P", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_S(2) = subGetChemValue(vSP, "S", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_CR(2) = subGetChemValue(vSP, "Cr", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_V(2) = subGetChemValue(vSP, "V", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_MO(2) = subGetChemValue(vSP, "Mo", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_CU(2) = subGetChemValue(vSP, "Cu", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_NI(2) = subGetChemValue(vSP, "Ni", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_B(2) = subGetChemValue(vSP, "B", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
        
        End If
        
'----------------------------------------------------------------------------------------------
        
        For i = 0 To 2
            
            Chem_Val(i) = subGetChemCeq(FOMULA_CD, CHEM_C(i), CHEM_SI(i), CHEM_MN(i), CHEM_P(i), CHEM_S(i), CHEM_CR(i), CHEM_V(i), _
                           CHEM_MO(i), CHEM_CU(i), CHEM_NI(i), CHEM_B(i))
            
        Next i
        
        .Row = iRow
        .Col = iSeq1: .Text = Chem_Val(0)
        .Col = iSeq2: .Text = Chem_Val(1)
        
        If sKnd = "3" Then
            .Col = iSeq3: .Text = Chem_Val(2)
        End If
    
    End With
    
End Sub



'---------------------------------------------------------------------------------------
'   1.ID           : subGetChemCeq
'   2.Name         : from GF_GetCeqValue Function -> call
'   3.Input  Value : Spread , CHEMICAL CODE , COL , ROW
'   4.Return Value : Chemical Code
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : from GS_SetChemicalLength call
'---------------------------------------------------------------------------------------
Private Function subGetChemValue(ByVal vSP As vaSpread, chem As String, ByVal iCode1 As Integer, ByVal iCode2 As Integer, ByVal iCode3 As Integer, ByVal iCol1 As Integer, ByVal iCol2 As Integer, ByVal iCol3 As Integer) As Double
    
    Dim i As Integer
    
    With vSP
        
        For i = 1 To .MaxRows
                            
            If Gf_Get_Cell_Value(vSP, i, iCode1) = chem Then
                subGetChemValue = Val(Gf_Get_Cell_Value(vSP, i, iCol1))
                Exit Function
            End If
        
        Next i
    
        
        For i = 1 To .MaxRows
                    
            If Gf_Get_Cell_Value(vSP, i, iCode2) = chem Then
                subGetChemValue = Val(Gf_Get_Cell_Value(vSP, i, iCol2))
                Exit Function
            End If
        
        Next i
        
        
        For i = 1 To .MaxRows
            
            If Gf_Get_Cell_Value(vSP, i, iCode3) = chem Then
                subGetChemValue = Val(Gf_Get_Cell_Value(vSP, i, iCol3))
                Exit Function
            End If
        
        Next i
        
    
    End With
    
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GS_SetChemicalLength
'   2.Name         : Chemical Code Length -> Column Apply
'   3.Input  Value : Spread ,Array Data , Col
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .06
'   7.Modify Date  :
'   8.Comment      : Chemical Code Length -> Column Apply
'---------------------------------------------------------------------------------------
Public Sub GS_SetChemicalLength(ByVal vSP As vaSpread, ByVal ArrayRecords As Variant, ByVal sCol As String, ByVal sKnd As String)

    Dim i As Integer
    Dim j As Integer
    Dim dblChem As Double
    Dim iCol(3) As Integer
    Dim iLoop As Integer
    
    
    iCol(1) = CInt(Mid(sCol, 1, 2))
    iCol(2) = CInt(Mid(sCol, 3, 2))
    iCol(3) = CInt(Mid(sCol, 5, 2))
    
    With vSP
        
        For iLoop = 1 To 3
                        
                For i = 1 To .MaxRows
                   
                   .Col = iCol(iLoop): .Row = i
                    
                    For j = 0 To UBound(ArrayRecords, 2)
                        If ArrayRecords(1, j) = Trim(.Text) Then
                            dblChem = ArrayRecords(2, j)
                            Exit For
                        End If
                    Next j
                
                    Call subSetChemLength(vSP, dblChem, .Col, .Row, sKnd)
                
                Next i
        
        Next iLoop
                
    End With

End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GS_CeqColumnLock
'   2.Name         : Ceq Column -> Lock
'   3.Input  Value : Spread , 分类 Code
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : Ceq Column -> Lock
'---------------------------------------------------------------------------------------
Public Sub GS_CeqColumnLock(ByVal vSP As vaSpread, ByVal sKnd As String, ByVal sCol As String)

    Dim i As Integer
    Dim iCol(3) As Integer
    
    iCol(1) = CInt(Mid(sCol, 1, 2))
    iCol(2) = CInt(Mid(sCol, 3, 2))
    iCol(3) = CInt(Mid(sCol, 5, 2))

    With vSP
    
        If sKnd = "3" Then
    
            For i = 1 To .MaxRows
                .Row = i
                
                .Col = iCol(1): .Lock = False
                .Col = iCol(2): .Lock = False
                .Col = iCol(3): .Lock = False
                
            Next i
        
        Else
        
            For i = 1 To .MaxRows
                .Row = i
                
                .Col = iCol(1): .Text = "": .Lock = True
                .Col = iCol(1): .Text = "": .Lock = True
                .Col = iCol(1): .Text = "": .Lock = True
                
            Next i
        
        End If
        
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GS_CeqColumnLock
'   2.Name         : Chemical Spread - Line Color Setting
'   3.Input  Value : Spread
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : Chemical Spread - Line Color Setting
'---------------------------------------------------------------------------------------
Public Sub GS_SetChemicalSpreadLineColor(ByVal vSP As vaSpread, ByVal sCol As String)
    
    Dim i As Integer
    Dim iCol(2) As Integer
    
    iCol(1) = CInt(Mid(sCol, 1, 2))
    iCol(2) = CInt(Mid(sCol, 3, 2))
    
    With vSP
        
        .Col = iCol(1)
        
        For i = 1 To .MaxRows
            .Row = i
            .BackColor = &HE1E4CD
        Next i
        
        .Col = iCol(2)
        For i = 1 To .MaxRows
            .Row = i
            .BackColor = &HE1E4CD
        Next i
    
    End With
    
End Sub




'---------------------------------------------------------------------------------------
'   1.ID           : subGetChemCeq
'   2.Name         : from GF_GetCeqValue Function -> call
'   3.Input  Value : Spread , CHEMICAL CODE , COL , ROW
'   4.Return Value : Ceq Value
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : from GS_SetChemicalLength call
'---------------------------------------------------------------------------------------
Private Function subGetChemCeq(ByVal FOMULA_CD As String, ByVal CHEM_C As Double, ByVal CHEM_SI As Double, ByVal CHEM_MN As Double, _
                          ByVal CHEM_P As Double, ByVal CHEM_S As Double, ByVal CHEM_CR As Double, ByVal CHEM_V As Double, _
                          ByVal CHEM_MO As Double, ByVal CHEM_CU As Double, ByVal CHEM_NI As Double, ByVal CHEM_B As Double) As Double
                          
    Dim V_CEQ As Double
           
    If FOMULA_CD = "A" Then
          V_CEQ = CHEM_C + CHEM_MN / 6
    
    ElseIf FOMULA_CD = "B" Then
          V_CEQ = CHEM_C + CHEM_MN / 6 + (CHEM_CR + CHEM_V + CHEM_MO) / 5 + (CHEM_CU + CHEM_NI) / 15
                                
    
    ElseIf FOMULA_CD = "C" Then
          V_CEQ = CHEM_C + CHEM_MN / 6 + (CHEM_SI / 24) + (CHEM_CR / 5) + (CHEM_MO / 4) + (CHEM_V / 14)
    
    
    ElseIf FOMULA_CD = "D" Then
          V_CEQ = CHEM_C + CHEM_MN / 6 + (CHEM_CU / 40) + (CHEM_NI / 20) + (CHEM_CR / 10) + (CHEM_MO / 50) - (CHEM_V / 10)
    
    
    ElseIf FOMULA_CD = "E" Then
          V_CEQ = CHEM_C + CHEM_SI / 30 + (CHEM_MN / 20) + (CHEM_CU / 20) + (CHEM_CR / 20) + (CHEM_NI / 60) + (CHEM_MO / 15) + (CHEM_V / 10) + (CHEM_B * 5)
    
    Else
          subGetChemCeq = 0
    End If
               
    subGetChemCeq = Round(V_CEQ, 2)

End Function


'---------------------------------------------------------------------------------------
'   1.ID           : subSetChemLength
'   2.Name         : from GS_SetChemicalLength Function -> call
'   3.Input  Value : Spread , CHEMICAL CODE , COL , ROW
'   4.Return Value : None
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : from GS_SetChemicalLength Function -> call
'---------------------------------------------------------------------------------------
Public Sub subSetChemLength(ByVal vSP As vaSpread, ByVal dblChem As Double, ByVal Col As Long, ByVal Row As Long, ByVal sKnd As String)

    With vSP
        
        .Row = Row
        
        .Col = Col + 1
        .TypeNumberMin = 0
        .TypeNumberMax = dblChem
        .TypeNumberDecPlaces = GF_GET_SPREAD_DECIMAL(Trim(str(dblChem)))
        
        
        
        .Col = Col + 2
        .TypeNumberMin = 0
        .TypeNumberMax = dblChem
        .TypeNumberDecPlaces = GF_GET_SPREAD_DECIMAL(Trim(str(dblChem)))
        
        
        If sKnd = "3" Then
        
            .Col = Col + 3
            .TypeNumberMin = 0
            .TypeNumberMax = dblChem
            .TypeNumberDecPlaces = GF_GET_SPREAD_DECIMAL(Trim(str(dblChem)))
            
        End If
    
    End With

End Sub


'#######################################################################################################################################################
'########################################################## Matrial INPUT CHECK ################################################################################
'#######################################################################################################################################################


'---------------------------------------------------------------------------------------
'   1.ID           : GF_COMMON_INPUT_CHECK
'   2.Name         : 材质 COMMON Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 COMMON Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_COMMON_INPUT_CHECK(oInput1 As Object, oInput2 As Object, Optional oInput3 As Object, Optional oInput4 As Object, Optional oInput5 As Object, Optional oInput6 As Object) As Boolean

    Dim iDataCnt As Integer
    Dim iChk As Integer
    
    oInput1.BackColor = vbWhite
    oInput2.BackColor = vbWhite
    
    iChk = 2
    
    If subObjectValue(oInput1) <> "" Then iDataCnt = iDataCnt + 1
    If subObjectValue(oInput2) <> "" Then iDataCnt = iDataCnt + 1
    
    If Not oInput3 Is Nothing Then
        oInput3.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput3) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If Not oInput4 Is Nothing Then
        oInput4.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput4) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If Not oInput5 Is Nothing Then
        oInput5.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput5) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If Not oInput6 Is Nothing Then
        oInput6.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput6) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If iDataCnt = 0 Or iDataCnt = iChk Then GoTo ResultTrue
    
    If subObjectValue(oInput1) = "" Then oInput1.BackColor = &HC0E0FF
    If subObjectValue(oInput2) = "" Then oInput2.BackColor = &HC0E0FF
    
    If Not oInput3 Is Nothing Then
        If subObjectValue(oInput3) = "" Then oInput3.BackColor = &HC0E0FF
    End If
    
    If Not oInput4 Is Nothing Then
        If subObjectValue(oInput4) = "" Then oInput4.BackColor = &HC0E0FF
    End If
    
    If Not oInput5 Is Nothing Then
        If subObjectValue(oInput5) = "" Then oInput5.BackColor = &HC0E0FF
    End If
    
    If Not oInput6 Is Nothing Then
        If subObjectValue(oInput6) = "" Then oInput6.BackColor = &HC0E0FF
    End If
                 
    Exit Function

ResultTrue:
        GF_MATR_COMMON_INPUT_CHECK = True
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_MIN_MAX_INPUT_CHECK
'   2.Name         : 材质 MIN , MAX Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 MIN , MAX Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_MIN_MAX_INPUT_CHECK(ByVal oMin As Object, ByVal oMax As Object, Optional ByVal oInput1 As Object, Optional ByVal oInput2 As Object, Optional ByVal oInput3 As Object, Optional ByVal oInput4 As Object, Optional ByVal oInput5 As Object) As Boolean

    Dim iDataCnt As Integer
    Dim iChk As Integer
    Dim bCheck As Boolean
    
    oMin.BackColor = vbWhite
    oMax.BackColor = vbWhite
    
    iChk = 1
    
    If subObjectValue(oMin) <> "" Or subObjectValue(oMax) <> "" Then iDataCnt = iDataCnt + 1
    
    If Not oInput1 Is Nothing Then
        oInput1.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput1) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If Not oInput2 Is Nothing Then
        oInput2.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput2) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If Not oInput3 Is Nothing Then
        oInput3.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput3) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If Not oInput4 Is Nothing Then
        oInput4.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput4) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    If Not oInput5 Is Nothing Then
        oInput5.BackColor = vbWhite: iChk = iChk + 1
        If subObjectValue(oInput5) <> "" Then iDataCnt = iDataCnt + 1
    End If
    
    
    bCheck = True
                        
    If subObjectValue(oMin) <> "" And subObjectValue(oMax) <> "" Then
        
        If oMin.Value >= oMax.Value Then bCheck = False
    
    ElseIf subObjectValue(oMin) = "" And subObjectValue(oMax) = "" Then
        
        bCheck = False
    
    End If
    
    If iDataCnt = 0 Or (iDataCnt = iChk And bCheck = True) Then GoTo ResultTrue
    
    If bCheck = False Then
        oMin.BackColor = &HC0E0FF
        oMax.BackColor = &HC0E0FF
    End If
    
    If Not oInput1 Is Nothing Then
        If subObjectValue(oInput1) = "" Then oInput1.BackColor = &HC0E0FF
    End If
    
    If Not oInput2 Is Nothing Then
        If subObjectValue(oInput2) = "" Then oInput2.BackColor = &HC0E0FF
    End If
    
    If Not oInput3 Is Nothing Then
        If subObjectValue(oInput3) = "" Then oInput3.BackColor = &HC0E0FF
    End If
    
    If Not oInput4 Is Nothing Then
        If subObjectValue(oInput4) = "" Then oInput4.BackColor = &HC0E0FF
    End If
                     
    If Not oInput5 Is Nothing Then
        If subObjectValue(oInput5) = "" Then oInput5.BackColor = &HC0E0FF
    End If
                                 
    Exit Function

ResultTrue:
        GF_MATR_MIN_MAX_INPUT_CHECK = True
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_IMPACT_INPUT_CHECK
'   2.Name         : 材质 - 冲击试验 Input Check
'   3.Input  Value : IMPACT_SMP_CD,IMPACT_KND,IMPACT_DIR,IMPACT_MIN,IMPACT_AVE_MIN,IMPACT_RATE_MIN,IMPACT_RATE_MAX,IMPACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 - 冲击试验 Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_IMPACT_INPUT_CHECK(ByVal txt_IMPACT_SMP_CD As Object, ByVal txt_IMPACT_KND As Object, ByVal txt_IMPACT_DIR As Object, ByVal sdb_IMPACT_MIN As Object, sdb_IMPACT_MIN_MIN As Object, ByVal sdb_IMPACT_AVE_MIN As Object, ByVal sdb_IMPACT_RATE_MIN As Object, ByVal sdb_IMPACT_RATE_MAX As Object, ByVal txt_IMPACT_DSC_CD As Object) As Boolean

    Dim iCnt As Integer
    Dim iChk As Integer
    Dim sMsg As String
    Dim bCheck As Boolean
    
    txt_IMPACT_SMP_CD.BackColor = vbWhite
    txt_IMPACT_KND.BackColor = vbWhite
    txt_IMPACT_DIR.BackColor = vbWhite
    sdb_IMPACT_MIN.BackColor = vbWhite
    sdb_IMPACT_MIN_MIN.BackColor = vbWhite
    sdb_IMPACT_AVE_MIN.BackColor = vbWhite
    sdb_IMPACT_RATE_MIN.BackColor = vbWhite
    sdb_IMPACT_RATE_MAX.BackColor = vbWhite
    txt_IMPACT_DSC_CD.BackColor = vbWhite
    
    bCheck = True
        
    If txt_IMPACT_SMP_CD.Text <> "" Then iCnt = iCnt + 1
    
    If txt_IMPACT_KND.Text <> "" Then iCnt = iCnt + 1
    
    If txt_IMPACT_DIR.Text <> "" Then iCnt = iCnt + 1
    
    
    If sdb_IMPACT_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_IMPACT_MIN_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_IMPACT_AVE_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_IMPACT_RATE_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_IMPACT_RATE_MAX.Value <> 0 Then iChk = iChk + 1
    
    If sdb_IMPACT_RATE_MIN.Value <> 0 And sdb_IMPACT_RATE_MAX.Value <> 0 Then
        If sdb_IMPACT_RATE_MIN.Value >= sdb_IMPACT_RATE_MAX.Value Then
            bCheck = False
            sdb_IMPACT_RATE_MIN.BackColor = &HC0E0FF
            sdb_IMPACT_RATE_MAX.BackColor = &HC0E0FF
        End If
    End If
                
    If iChk > 0 Then iCnt = iCnt + 1
            
    If txt_IMPACT_DSC_CD.Text <> "" Then iCnt = iCnt + 1
            
    If iCnt = 0 Or (iCnt = 5 And bCheck = True) Then GoTo ResultTrue
    
    If txt_IMPACT_SMP_CD.Text = "" Then
        txt_IMPACT_SMP_CD.BackColor = &HC0E0FF
        sMsg = sMsg + "冲击试验 - 取样代码必须输入" & Chr(13) & Chr(10)
    End If
    
    If txt_IMPACT_KND.Text = "" Then
        txt_IMPACT_KND.BackColor = &HC0E0FF
        sMsg = sMsg + "冲击试验 - 缺口类型代码必须输入" & Chr(13) & Chr(10)
    End If
    
    If txt_IMPACT_DIR.Text = "" Then
        txt_IMPACT_DIR.BackColor = &HC0E0FF
        sMsg = sMsg + "冲击试验 - 冲击方向代码必须输入" & Chr(13) & Chr(10)
    End If
    
    If iChk = 0 Then
        sdb_IMPACT_MIN.BackColor = &HC0E0FF
        sMsg = sMsg + "冲击试验 - 下限必须输入" & Chr(13) & Chr(10)
    End If
            
    If txt_IMPACT_DSC_CD.Text = "" Then
        txt_IMPACT_DSC_CD.BackColor = &HC0E0FF
        sMsg = sMsg + "冲击试验 - 判定代码必须输入" & Chr(13) & Chr(10)
    End If
        
   'If sMsg <> "" Then MsgBox sMsg
        
    Exit Function

ResultTrue:
        GF_MATR_IMPACT_INPUT_CHECK = True
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_TIM_IMPACT_INPUT_CHECK
'   2.Name         : 材质 - 时效冲击试验 Input Check
'   3.Input  Value : TIM_IMPACT_SMP_CD,TIM_IMPACT_KND,TIM_IMPACT_DIR,TIM_IMPACT_MIN,TIM_IMPACT_AVE_MIN,TIM_IMPACT_RATE_MIN,TIM_IMPACT_RATE_MAX,TIM_IMPACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 - 时效冲击试验 Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_TIM_IMPACT_INPUT_CHECK(ByVal txt_TIM_IMPACT_SMP_CD As Object, ByVal txt_TIM_IMPACT_KND As Object, ByVal txt_TIM_IMPACT_DIR As Object, ByVal sdb_TIM_IMPACT_MIN As Object, ByVal sdb_TIM_IMPACT_MIN_MIN As Object, ByVal sdb_TIM_IMPACT_AVE_MIN As Object, ByVal sdb_TIM_IMPACT_RATE_MIN As Object, ByVal sdb_TIM_IMPACT_RATE_MAX As Object, ByVal txt_TIM_IMPACT_DSC_CD As Object) As Boolean

    Dim iCnt As Integer
    Dim iChk As Integer
    Dim sMsg As String
    Dim bCheck As Boolean
    
    txt_TIM_IMPACT_SMP_CD.BackColor = vbWhite
    txt_TIM_IMPACT_KND.BackColor = vbWhite
    txt_TIM_IMPACT_DIR.BackColor = vbWhite
'    sdb_TIM_IMPACT_TIM.BackColor = vbWhite
    sdb_TIM_IMPACT_MIN.BackColor = vbWhite
    sdb_TIM_IMPACT_MIN_MIN.BackColor = vbWhite
    sdb_TIM_IMPACT_AVE_MIN.BackColor = vbWhite
    sdb_TIM_IMPACT_RATE_MIN.BackColor = vbWhite
    sdb_TIM_IMPACT_RATE_MAX.BackColor = vbWhite
    txt_TIM_IMPACT_DSC_CD.BackColor = vbWhite
    
    bCheck = True
        
    If txt_TIM_IMPACT_SMP_CD.Text <> "" Then iCnt = iCnt + 1
    
    If txt_TIM_IMPACT_KND.Text <> "" Then iCnt = iCnt + 1
    
    If txt_TIM_IMPACT_DIR.Text <> "" Then iCnt = iCnt + 1
    
'    If sdb_TIM_IMPACT_TIM.Value <> 0 Then iCnt = iCnt + 1
    
    
    If sdb_TIM_IMPACT_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_TIM_IMPACT_MIN_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_TIM_IMPACT_AVE_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_TIM_IMPACT_RATE_MIN.Value <> 0 Then iChk = iChk + 1
    If sdb_TIM_IMPACT_RATE_MAX.Value <> 0 Then iChk = iChk + 1
    
    If sdb_TIM_IMPACT_RATE_MIN.Value <> 0 And sdb_TIM_IMPACT_RATE_MAX.Value <> 0 Then
        If sdb_TIM_IMPACT_RATE_MIN.Value >= sdb_TIM_IMPACT_RATE_MAX.Value Then
            bCheck = False
            sdb_TIM_IMPACT_RATE_MIN.BackColor = &HC0E0FF
            sdb_TIM_IMPACT_RATE_MAX.BackColor = &HC0E0FF
        End If
    End If
                
    If iChk > 0 Then iCnt = iCnt + 1
            
    If txt_TIM_IMPACT_DSC_CD.Text <> "" Then iCnt = iCnt + 1
            
    If iCnt = 0 Or (iCnt = 5 And bCheck = True) Then GoTo ResultTrue
    
    If txt_TIM_IMPACT_SMP_CD.Text = "" Then
        txt_TIM_IMPACT_SMP_CD.BackColor = &HC0E0FF
        sMsg = sMsg + "时效冲击试验 - 取样代码必须输入" & Chr(13) & Chr(10)
    End If
    
    If txt_TIM_IMPACT_KND.Text = "" Then
        txt_TIM_IMPACT_KND.BackColor = &HC0E0FF
        sMsg = sMsg + "时效冲击试验 - 缺口类型代码必须输入" & Chr(13) & Chr(10)
    End If
    
    If txt_TIM_IMPACT_DIR.Text = "" Then
        txt_TIM_IMPACT_DIR.BackColor = &HC0E0FF
        sMsg = sMsg + "时效冲击试验 - 冲击方向代码必须输入" & Chr(13) & Chr(10)
    End If
    
'    If sdb_TIM_IMPACT_TIM.Value = 0 Then
'        sdb_TIM_IMPACT_TIM.BackColor = &HC0E0FF
'        sMsg = sMsg + "时效冲击试验 - 时效时间必须输入" & Chr(13) & Chr(10)
'    End If
        
    If iChk = 0 Then
        sdb_TIM_IMPACT_MIN.BackColor = &HC0E0FF
        sMsg = sMsg + "时效冲击试验 - 下限值必须输入" & Chr(13) & Chr(10)
    End If
            
    If txt_TIM_IMPACT_DSC_CD.Text = "" Then
        txt_TIM_IMPACT_DSC_CD.BackColor = &HC0E0FF
        sMsg = sMsg + "时效冲击试验 - 判定代码必须输入" & Chr(13) & Chr(10)
    End If
        
   'If sMsg <> "" Then MsgBox sMsg
        
    Exit Function

ResultTrue:
        GF_MATR_TIM_IMPACT_INPUT_CHECK = True
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_HIC_INPUT_CHECK
'   2.Name         : 材质 - 抗氢裂能力 Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 - 抗氢裂能力 Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_HIC_INPUT_CHECK(ByVal oInput1 As Object, ByVal oInput2 As Object, ByVal oInput3 As Object, ByVal oInput4 As Object, ByVal oInput5 As Object, ByVal oInput6 As Object) As Boolean

    Dim iCnt As Integer
    
    oInput1.BackColor = vbWhite
    oInput2.BackColor = vbWhite
    oInput3.BackColor = vbWhite
    oInput4.BackColor = vbWhite
    oInput5.BackColor = vbWhite
    oInput6.BackColor = vbWhite
           
    If subObjectValue(oInput1) <> "" Then iCnt = iCnt + 1
    If subObjectValue(oInput2) <> "" Then iCnt = iCnt + 1
    If subObjectValue(oInput3) <> "" Or subObjectValue(oInput4) <> "" Or subObjectValue(oInput5) <> "" Then iCnt = iCnt + 1
    
    If subObjectValue(oInput6) <> "" Then iCnt = iCnt + 1
            
    If iCnt = 0 Or iCnt >= 4 Then GoTo ResultTrue
    
    If subObjectValue(oInput1) = "" Then oInput1.BackColor = &HC0E0FF
    If subObjectValue(oInput2) = "" Then oInput2.BackColor = &HC0E0FF
    
    If subObjectValue(oInput3) = "" And subObjectValue(oInput4) = "" And subObjectValue(oInput5) = "" Then
        oInput3.BackColor = &HC0E0FF
        oInput4.BackColor = &HC0E0FF
        oInput5.BackColor = &HC0E0FF
    End If
        
    If subObjectValue(oInput6) = "" Then oInput6.BackColor = &HC0E0FF
        
    Exit Function

ResultTrue:
        GF_MATR_HIC_INPUT_CHECK = True
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_ACD_DFT_INPUT_CHECK
'   2.Name         : 材质 - 酸浸检验 Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 - 酸浸检验 Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_ACD_DFT_INPUT_CHECK(ByVal oTyp1 As Object, ByVal oGrd1 As Object, ByVal oTyp2 As Object, ByVal oGrd2 As Object, ByVal oTyp3 As Object, ByVal oGrd3 As Object, ByVal oTyp4 As Object, ByVal oGrd4 As Object, ByVal oTyp5 As Object, ByVal oGrd5 As Object, ByVal oDsc As Object) As Boolean

    Dim iCnt As Integer
    Dim iChk As Integer
    Dim bChk1 As Boolean
    Dim bChk2 As Boolean
    Dim bChk3 As Boolean
    Dim bChk4 As Boolean
    Dim bChk5 As Boolean
    Dim bCheck As Boolean
    
    oTyp1.BackColor = vbWhite
    oTyp2.BackColor = vbWhite
    oTyp3.BackColor = vbWhite
    oTyp4.BackColor = vbWhite
    oTyp5.BackColor = vbWhite
    oGrd1.BackColor = vbWhite
    oGrd2.BackColor = vbWhite
    oGrd3.BackColor = vbWhite
    oGrd4.BackColor = vbWhite
    oGrd5.BackColor = vbWhite
    oDsc.BackColor = vbWhite
           
    If subObjectValue(oTyp1) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp2) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp3) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp4) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp5) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd1) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd2) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd3) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd4) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd5) <> "" Then iChk = iChk + 1
    
    iCnt = iChk
    
    If oDsc.Text <> "" Then iCnt = iCnt + 1
                      
    If (subObjectValue(oTyp1) <> "" And subObjectValue(oGrd1) <> "") Or (subObjectValue(oTyp1) = "" And subObjectValue(oGrd1) = "") Then bChk1 = True
    If (subObjectValue(oTyp2) <> "" And subObjectValue(oGrd2) <> "") Or (subObjectValue(oTyp2) = "" And subObjectValue(oGrd2) = "") Then bChk2 = True
    If (subObjectValue(oTyp3) <> "" And subObjectValue(oGrd3) <> "") Or (subObjectValue(oTyp3) = "" And subObjectValue(oGrd3) = "") Then bChk3 = True
    If (subObjectValue(oTyp4) <> "" And subObjectValue(oGrd4) <> "") Or (subObjectValue(oTyp4) = "" And subObjectValue(oGrd4) = "") Then bChk4 = True
    If (subObjectValue(oTyp5) <> "" And subObjectValue(oGrd5) <> "") Or (subObjectValue(oTyp5) = "" And subObjectValue(oGrd5) = "") Then bChk5 = True

    If (bChk1 = True And bChk2 = True And bChk3 = True And bChk4 = True And bChk5 = True) And oDsc.Text <> "" And iChk > 0 Then bCheck = True
            
    If iCnt = 0 Or bCheck = True Then GoTo ResultTrue
    
    
    If bChk1 = False And (oTyp1.Text <> "" Or subObjectValue(oGrd1) <> "") Then
        oTyp1.BackColor = &HC0E0FF
        oGrd1.BackColor = &HC0E0FF
    End If
    
    If bChk2 = False And (oTyp2.Text <> "" Or subObjectValue(oGrd2) <> "") Then
        oTyp2.BackColor = &HC0E0FF
        oGrd2.BackColor = &HC0E0FF
    End If
    
    If bChk3 = False And (oTyp3.Text <> "" Or subObjectValue(oGrd3) <> "") Then
        oTyp3.BackColor = &HC0E0FF
        oGrd3.BackColor = &HC0E0FF
    End If
    
    If bChk4 = False And (oTyp4.Text <> "" Or subObjectValue(oGrd4) <> "") Then
        oTyp4.BackColor = &HC0E0FF
        oGrd4.BackColor = &HC0E0FF
    End If
    
    If bChk5 = False And (oTyp5.Text <> "" Or subObjectValue(oGrd5) <> "") Then
        oTyp5.BackColor = &HC0E0FF
        oGrd5.BackColor = &HC0E0FF
    End If

    If oDsc.Text = "" Then
        oDsc.BackColor = &HC0E0FF
    ElseIf iChk = 0 Then
        oTyp1.BackColor = &HC0E0FF
        oGrd1.BackColor = &HC0E0FF
    End If
        
    Exit Function

ResultTrue:
        GF_MATR_ACD_DFT_INPUT_CHECK = True
End Function



'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_FRACT_INPUT_CHECK
'   2.Name         : 材质 - 断口检验 Input Check
'   3.Input  Value : FRACT_SMP_CD , FRACT_KND , txt_FRACT_NAME_CD1 ,txt_FRACT_GRD1 , txt_FRACT_NAME_CD2 , txt_FRACT_GRD2 , txt_FRACT_NAME_CD3 ,txt_FRACT_GRD3 , txt_FRACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 - 断口检验 Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_FRACT_INPUT_CHECK(ByVal txt_FRACT_SMP_CD As Object, ByVal oTyp1 As Object, ByVal oGrd1 As Object, ByVal oTyp2 As Object, ByVal oGrd2 As Object, ByVal oTyp3 As Object, ByVal oGrd3 As Object, ByVal oTyp4 As Object, ByVal oGrd4 As Object, ByVal oTyp5 As Object, ByVal oGrd5 As Object, ByVal oDsc As Object) As Boolean


    Dim iCnt As Integer
    Dim iChk As Integer
    Dim bChk1 As Boolean
    Dim bChk2 As Boolean
    Dim bChk3 As Boolean
    Dim bChk4 As Boolean
    Dim bChk5 As Boolean
    Dim bCheck As Boolean
    
    oTyp1.BackColor = vbWhite
    oTyp2.BackColor = vbWhite
    oTyp3.BackColor = vbWhite
    oTyp4.BackColor = vbWhite
    oTyp5.BackColor = vbWhite
    oGrd1.BackColor = vbWhite
    oGrd2.BackColor = vbWhite
    oGrd3.BackColor = vbWhite
    oGrd4.BackColor = vbWhite
    oGrd5.BackColor = vbWhite
    oDsc.BackColor = vbWhite
           
    If subObjectValue(oTyp1) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp2) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp3) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp4) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp5) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd1) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd2) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd3) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd4) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd5) <> "" Then iChk = iChk + 1
    
    iCnt = iChk
    
    If txt_FRACT_SMP_CD.Text <> "" Then iCnt = iCnt + 1
    If oDsc.Text <> "" Then iCnt = iCnt + 1
                      
    If (subObjectValue(oTyp1) <> "" And subObjectValue(oGrd1) <> "") Or (subObjectValue(oTyp1) = "" And subObjectValue(oGrd1) = "") Then bChk1 = True
    If (subObjectValue(oTyp2) <> "" And subObjectValue(oGrd2) <> "") Or (subObjectValue(oTyp2) = "" And subObjectValue(oGrd2) = "") Then bChk2 = True
    If (subObjectValue(oTyp3) <> "" And subObjectValue(oGrd3) <> "") Or (subObjectValue(oTyp3) = "" And subObjectValue(oGrd3) = "") Then bChk3 = True
    If (subObjectValue(oTyp4) <> "" And subObjectValue(oGrd4) <> "") Or (subObjectValue(oTyp4) = "" And subObjectValue(oGrd4) = "") Then bChk4 = True
    If (subObjectValue(oTyp5) <> "" And subObjectValue(oGrd5) <> "") Or (subObjectValue(oTyp5) = "" And subObjectValue(oGrd5) = "") Then bChk5 = True

    If (bChk1 = True And bChk2 = True And bChk3 = True And bChk4 = True And bChk5 = True) And oDsc.Text <> "" And iChk > 0 Then bCheck = True
            
    If iCnt = 0 Or bCheck = True Then GoTo ResultTrue
    
    
    If bChk1 = False And (oTyp1.Text <> "" Or subObjectValue(oGrd1) <> "") Then
        oTyp1.BackColor = &HC0E0FF
        oGrd1.BackColor = &HC0E0FF
    End If
    
    If bChk2 = False And (oTyp2.Text <> "" Or subObjectValue(oGrd2) <> "") Then
        oTyp2.BackColor = &HC0E0FF
        oGrd2.BackColor = &HC0E0FF
    End If
    
    If bChk3 = False And (oTyp3.Text <> "" Or subObjectValue(oGrd3) <> "") Then
        oTyp3.BackColor = &HC0E0FF
        oGrd3.BackColor = &HC0E0FF
    End If
    
    If bChk4 = False And (oTyp4.Text <> "" Or subObjectValue(oGrd4) <> "") Then
        oTyp4.BackColor = &HC0E0FF
        oGrd4.BackColor = &HC0E0FF
    End If
    
    If bChk5 = False And (oTyp5.Text <> "" Or subObjectValue(oGrd5) <> "") Then
        oTyp5.BackColor = &HC0E0FF
        oGrd5.BackColor = &HC0E0FF
    End If

    If oDsc.Text = "" Then
        oDsc.BackColor = &HC0E0FF
    ElseIf iChk = 0 Then
        oTyp1.BackColor = &HC0E0FF
        oGrd1.BackColor = &HC0E0FF
    End If
        
    Exit Function

ResultTrue:
        GF_MATR_FRACT_INPUT_CHECK = True
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_NON_METAL_INPUT_CHECK
'   2.Name         : 材质 - 非金属夹杂 Input Check
'   3.Input  Value : FRACT_SMP_CD , FRACT_KND , txt_FRACT_NAME_CD1 ,txt_FRACT_GRD1 , txt_FRACT_NAME_CD2 , txt_FRACT_GRD2 , txt_FRACT_NAME_CD3 ,txt_FRACT_GRD3 , txt_FRACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : 材质 - 非金属夹杂 Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_NON_METAL_INPUT_CHECK(ByVal txt_NON_METAL_SMP_CD As Object, ByVal txt_NON_METAL_TYP As TextBox, ByVal oTyp1 As Object, ByVal oGrd1 As Object, ByVal oTyp2 As Object, ByVal oGrd2 As Object, ByVal oTyp3 As Object, ByVal oGrd3 As Object, ByVal oTyp4 As Object, ByVal oGrd4 As Object, ByVal oTyp5 As Object, ByVal oGrd5 As Object, ByVal oTyp6 As Object, ByVal oGrd6 As Object, ByVal oTyp7 As Object, ByVal oGrd7 As Object, ByVal oTyp8 As Object, ByVal oGrd8 As Object, ByVal oDsc As Object) As Boolean

    Dim iCnt As Integer
    Dim iChk As Integer
    Dim bChk1 As Boolean
    Dim bChk2 As Boolean
    Dim bChk3 As Boolean
    Dim bChk4 As Boolean
    Dim bChk5 As Boolean
    Dim bChk6 As Boolean
    Dim bChk7 As Boolean
    Dim bChk8 As Boolean
    Dim bCheck As Boolean
    
    oTyp1.BackColor = vbWhite
    oTyp2.BackColor = vbWhite
    oTyp3.BackColor = vbWhite
    oTyp4.BackColor = vbWhite
    oTyp5.BackColor = vbWhite
    oTyp6.BackColor = vbWhite
    oTyp7.BackColor = vbWhite
    oTyp8.BackColor = vbWhite
    
    oGrd1.BackColor = vbWhite
    oGrd2.BackColor = vbWhite
    oGrd3.BackColor = vbWhite
    oGrd4.BackColor = vbWhite
    oGrd5.BackColor = vbWhite
    oGrd6.BackColor = vbWhite
    oGrd7.BackColor = vbWhite
    oGrd8.BackColor = vbWhite
    oDsc.BackColor = vbWhite
           
    If subObjectValue(oTyp1) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp2) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp3) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp4) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp5) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp6) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp7) <> "" Then iChk = iChk + 1
    If subObjectValue(oTyp8) <> "" Then iChk = iChk + 1
    
    If subObjectValue(oGrd1) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd2) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd3) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd4) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd5) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd6) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd7) <> "" Then iChk = iChk + 1
    If subObjectValue(oGrd8) <> "" Then iChk = iChk + 1
        
    iCnt = iChk
    
    If txt_NON_METAL_SMP_CD.Text <> "" Then iCnt = iCnt + 1
    If txt_NON_METAL_TYP.Text <> "" Then iCnt = iCnt + 1
    If oDsc.Text <> "" Then iCnt = iCnt + 1
                      
    If (subObjectValue(oTyp1) <> "" And subObjectValue(oGrd1) <> "") Or (subObjectValue(oTyp1) = "" And subObjectValue(oGrd1) = "") Then bChk1 = True
    If (subObjectValue(oTyp2) <> "" And subObjectValue(oGrd2) <> "") Or (subObjectValue(oTyp2) = "" And subObjectValue(oGrd2) = "") Then bChk2 = True
    If (subObjectValue(oTyp3) <> "" And subObjectValue(oGrd3) <> "") Or (subObjectValue(oTyp3) = "" And subObjectValue(oGrd3) = "") Then bChk3 = True
    If (subObjectValue(oTyp4) <> "" And subObjectValue(oGrd4) <> "") Or (subObjectValue(oTyp4) = "" And subObjectValue(oGrd4) = "") Then bChk4 = True
    If (subObjectValue(oTyp5) <> "" And subObjectValue(oGrd5) <> "") Or (subObjectValue(oTyp5) = "" And subObjectValue(oGrd5) = "") Then bChk5 = True
    If (subObjectValue(oTyp6) <> "" And subObjectValue(oGrd6) <> "") Or (subObjectValue(oTyp6) = "" And subObjectValue(oGrd6) = "") Then bChk6 = True
    If (subObjectValue(oTyp7) <> "" And subObjectValue(oGrd7) <> "") Or (subObjectValue(oTyp7) = "" And subObjectValue(oGrd7) = "") Then bChk7 = True
    If (subObjectValue(oTyp8) <> "" And subObjectValue(oGrd8) <> "") Or (subObjectValue(oTyp8) = "" And subObjectValue(oGrd8) = "") Then bChk8 = True

    If (bChk1 = True And bChk2 = True And bChk3 = True And bChk4 = True And bChk5 = True And bChk6 = True And bChk7 = True And bChk8 = True) And oDsc.Text <> "" And iChk > 0 Then bCheck = True
            
    If iCnt = 0 Or bCheck = True Then GoTo ResultTrue
    
    If bChk1 = False And (oTyp1.Text <> "" Or subObjectValue(oGrd1) <> "") Then
        oTyp1.BackColor = &HC0E0FF
        oGrd1.BackColor = &HC0E0FF
    End If
    
    If bChk2 = False And (oTyp2.Text <> "" Or subObjectValue(oGrd2) <> "") Then
        oTyp2.BackColor = &HC0E0FF
        oGrd2.BackColor = &HC0E0FF
    End If
    
    If bChk3 = False And (oTyp3.Text <> "" Or subObjectValue(oGrd3) <> "") Then
        oTyp3.BackColor = &HC0E0FF
        oGrd3.BackColor = &HC0E0FF
    End If
    
    If bChk4 = False And (oTyp4.Text <> "" Or subObjectValue(oGrd4) <> "") Then
        oTyp4.BackColor = &HC0E0FF
        oGrd4.BackColor = &HC0E0FF
    End If
    
    If bChk5 = False And (oTyp5.Text <> "" Or subObjectValue(oGrd5) <> "") Then
        oTyp5.BackColor = &HC0E0FF
        oGrd5.BackColor = &HC0E0FF
    End If
    
    If bChk6 = False And (oTyp6.Text <> "" Or subObjectValue(oGrd6) <> "") Then
        oTyp6.BackColor = &HC0E0FF
        oGrd6.BackColor = &HC0E0FF
    End If
    
    If bChk7 = False And (oTyp7.Text <> "" Or subObjectValue(oGrd7) <> "") Then
        oTyp7.BackColor = &HC0E0FF
        oGrd7.BackColor = &HC0E0FF
    End If
    
    If bChk8 = False And (oTyp8.Text <> "" Or subObjectValue(oGrd8) <> "") Then
        oTyp8.BackColor = &HC0E0FF
        oGrd8.BackColor = &HC0E0FF
    End If
    
    If oDsc.Text = "" Then
        oDsc.BackColor = &HC0E0FF
    ElseIf iChk = 0 Then
        oTyp1.BackColor = &HC0E0FF
        oGrd1.BackColor = &HC0E0FF
    End If
        
    Exit Function

ResultTrue:
        GF_MATR_NON_METAL_INPUT_CHECK = True
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : subObjectValue
'   2.Name         : Control value Return
'   3.Input  Value : Control Name
'   4.Return Value : String
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .27
'   7.Modify Date  :
'   8.Comment      : Control value Return
'---------------------------------------------------------------------------------------
Private Function subObjectValue(ByVal oControl As Object) As String
    
    If TypeOf oControl Is TextBox Then
        
        subObjectValue = Trim(oControl.Text)
    
    ElseIf TypeOf oControl Is sidbEdit Then
        
        If oControl.Value = 0 Then
            subObjectValue = ""
        Else
            subObjectValue = str(oControl.Value)
        End If
    
    Else
        
        subObjectValue = Trim(oControl.Text)
        
    End If


'    Select Case Left(oControl.Name, 3)
'
'        Case "txt"
'            subObjectValue = Trim(oControl.Text)
'        Case "sdb"
'            If oControl.Value = 0 Then
'                subObjectValue = ""
'            Else
'                subObjectValue = Str(oControl.Value)
'            End If
'        Case Else
'
'    End Select

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_GET_SPREAD_DECIMAL
'   2.Name         : Spread Decimal Value Read
'   3.Input  Value : Control Name
'   4.Return Value : String
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .27
'   7.Modify Date  :
'   8.Comment      : Spread Decimal Value Read
'---------------------------------------------------------------------------------------
Public Function GF_GET_SPREAD_DECIMAL(ByVal sVal As String) As Long
    
    Dim iLen As Integer
    Dim j As Integer
    
    j = InStr(1, sVal, ".", 1)
    
    iLen = Len(sVal) - InStr(1, sVal, ".", 1)
    
    GF_GET_SPREAD_DECIMAL = iLen
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GP_SELECT_ROW
'   2.Name         : Spread Row Select
'   3.Input  Value : Spread Name , Spread Row
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .27
'   7.Modify Date  :
'   8.Comment      : Spread Row Select
'---------------------------------------------------------------------------------------
Public Sub GP_SELECT_ROW(ByVal ss As vaSpread, ByVal iRow As Long)
    
    With ss
        
        .AllowMultiBlocks = True
        .SetSelection 1, iRow, .MaxCols, iRow
        .SetFocus
        
        Call Gp_Sp_EvenRowBackcolor(ss)
        
        'Call GP_SetRowHeaderClear(vSP, iRow)
        
        .Row = iRow
                
    End With
    
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GP_ROW_PASTE
'   2.Name         : Spread Row Paste
'   3.Input  Value : Spread Name , Spread Row , Master Collection
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .27
'   7.Modify Date  :
'   8.Comment      : Spread Row Paste
'---------------------------------------------------------------------------------------
Public Sub GP_ROW_PASTE(Sc As Collection, ByVal iRow As Long, Optional MC As Collection)
    
    Dim i As Long
    
    Dim iRow2 As Long
    
    If Sc("Spread").MaxRows = 0 Then Exit Sub
    
    iRow2 = Sc("Spread").ActiveRow + 1
        
    With Sc.Item("Spread")
        
        Call Gp_Sp_InsertRow(Sc("Spread"), .ActiveRow)
        
        For i = 1 To .MaxCols
                    
           Call GP_SET_CELL_VALUE(Sc.Item("Spread"), iRow2, i, Gf_Get_Cell_Value(Sc.Item("Spread"), iRow, i))
                        
        Next i

        .RowHeight(.ActiveRow) = 12.54
        
        Call Gp_Sp_ActiveCell(Sc.Item("Spread"), IIf(Sc.Item("First") > 0, Sc.Item("First"), 1))
                
    End With
    
    If Not MC Is Nothing Then
    
        Call Gp_Ms_ControlLock(MC("pcontrol"), False)
        
    End If
    
    
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GP_ROW_CANCEL
'   2.Name         : Spread Row Cancel (Insert, Update, Delete)
'   3.Input  Value : Conn Connection, Sc Collection
'   4.Return Value :
'   5.Writer       :
'   6.Create Date  : 2003. 10 .01
'   7.Modify Date  :
'   8.Comment      : Spread Row Cancel (Insert, Update, Delete)
'---------------------------------------------------------------------------------------
Public Sub GP_ROW_CANCEL(Sc As Collection)

On Error GoTo SpreadCancel_Error

    Dim sQuery As String
    Dim i As Integer
    Dim iRow As Long

    With Sc
        
        Screen.MousePointer = vbHourglass
        
        .Item("Spread").ReDraw = False
        
        If .Item("Spread").MaxRows < 1 Or .Item("Spread").SelBlockRow < 1 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        iRow = .Item("Spread").Row
        
                    
            Select Case Trim(Gf_Sp_RcvData(.Item("Spread"), 0, iRow))
                
                Case "Input"
                    Call Gp_Sp_DeleteRow(.Item("Spread"), iRow)
                    Call Gp_Sp_EvenRowBackcolor(.Item("Spread"))

                Case "Delete"
                    Call Gp_Sp_SendData(.Item("Spread"), "", 0, iRow)
                    Call Gp_Sp_EvenRowBackcolor(.Item("Spread"))
                                        
                Case "Update"
                    sQuery = Gf_Sp_MakeQuery(.Item("Spread"), .Item("P-O"), "O", .Item("pColumn"), iRow)
                    Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, .Item("Spread"), iRow)
                    Call Gp_Sp_SendData(.Item("Spread"), "", 0, iRow)
                    Call Gp_Sp_EvenRowBackcolor(.Item("Spread"))
                                        
            End Select
            
        
        .Item("Spread").ReDraw = True
        
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Sub
    
SpreadCancel_Error:

    Screen.MousePointer = vbDefault
    
End Sub



'---------------------------------------------------------------------------------------
'   1.ID           : GF_TextChangeCase
'   2.Name         : Force TextBox Input to Upper or Lower Case
'   3.Input  Value : KeyAscii , UpperCase
'   4.Return Value : Integer
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .27
'   7.Modify Date  :
'   8.Comment      : Force TextBox Input to Upper or Lower Case
'---------------------------------------------------------------------------------------
Function GF_TextChangeCase(KeyAscii As Integer, UpperCase As Boolean) As Integer
    GF_TextChangeCase = KeyAscii
    If UpperCase Then
        'Force characters into upper case
        If KeyAscii > 96 And KeyAscii < 123 Then
            'Typed letter from "a-z", map it to "A-Z"
            GF_TextChangeCase = KeyAscii - 32
        End If
    Else
        'Force into lower case
        If KeyAscii > 64 And KeyAscii < 91 Then
            'Typed letter from "A-Z", map it to "a-z"
            GF_TextChangeCase = KeyAscii + 32
        End If
    End If
End Function




'Private Sub txtInput_KeyPress(KeyAscii As Integer)
'    KeyAscii = GF_TextChangeCase(KeyAscii, True) 'Force the textbox into upper case
'End Sub

'Highlight all text in a textbox when it gets focus
'Private Sub Text1_GotFocus()
'    'put this in the GotFocus event
'    Text1.SelStart = 0
'    Text1.SelLength = Len(Text1.Text)
'End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GP_ROW_BACKCOLOR
'   2.Name         : Spread Row Backcolor Setting
'   3.Input  Value : Spread Name , Spread Row
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .27
'   7.Modify Date  :
'   8.Comment      : Spread Row Backcolor Setting
'---------------------------------------------------------------------------------------
Public Sub GP_ROW_BACKCOLOR(ByVal vSP As vaSpread)
    
    With vSP
        
        '.SelBackColor = &HFFC0C0
        '.SelBackColor = RGB(160, 180, 240)
        '.SelBackColor = RGB(180, 200, 240)
        .SelForeColor = vbBlack
        '.GridColor = vbBlue
        '.ShadowDark = &HFFC0C0
        'Gp_Sp_BlockColor
    End With
    
End Sub



'*************************************************************************************
'* Function Name : NullCheck
'* Description   : 涝仿 Data狼 Null巩磊甫 ""肺 摹券 棺 器镐屈侥,
'*                 漂荐巩磊 力芭, 巩磊屈 鸥涝 函券
'* Parameter :
'*          Input :Rsdata As Variant                             -> 涝仿磊丰
'*                 DefaultValueCondtion As Variant               -> 磊丰啊 Null老版快 摹券 巩磊蔼
'*                 DataFormatTypeCondition As String             -> 器镐屈侥蔼
'*                 DatainSymbolTextAllDeleteCondition As String  -> 力芭 且 巩磊蔼
'*                 DataTypeSettingCondition As Integer           -> 函券 且 鸥涝狼 备盒 箭磊蔼
'*                 (DataType -> 屈侥篮  1-String, 2-Integer, 3-Long, 4-Double)
'*
'*          荤侩抗力      nullcheck("扁夯蔼","null矫 摹券蔼","器镐屈侥","漂荐扁龋力芭",鸥涝函券)
'*     ----------------------------------------------------------------------------
'*     2002/03/01     捞犁霖[00]    Initial Coding
'*************************************************************************************
Public Function NullCheck(Rsdata As Variant, _
                          Optional DefaultValueCondtion As Variant = "", _
                          Optional DataFormatTypeCondition As String = "", _
                          Optional DatainSymbolTextAllDeleteCondition As String = "", _
                          Optional DataTypeSettingCondition As Integer = 0) As Variant

Dim Dv As Variant                   '扁夯 汲沥 巩磊
Dim Df As String                    '单捞鸥 器镐 汲沥 巩磊
Dim Ds As String                    '单捞鸥郴狼 力芭 且 漂荐扁龋
Dim Dt As Variant                   '单捞鸥 函券侩 鸥涝备盒 箭磊

    Dv = DefaultValueCondtion
    Df = DataFormatTypeCondition
    Ds = DatainSymbolTextAllDeleteCondition
    Dt = DataTypeSettingCondition
    
    'Null巩磊 备盒
    NullCheck = IIf(IsNull(Rsdata) Or Rsdata = "", Dv, Trim(Rsdata))
    
    '单捞鸥狼 荤侩磊 沥狼 扁龋 力芭
    If Ds <> "" Then
        NullCheck = SymbolTxtpart2(NullCheck, Ds)
    End If
    
    '单捞鸥 荤侩磊 沥狼 屈侥阑 Formating
    If Df <> "" Then
        NullCheck = Trim(NullCheck)
        NullCheck = Format(NullCheck, Df)
    End If
    
    '巩磊凯狼 DataType阑 函券窃
    'DataType -> 屈侥篮  1-String, 2-Integer, 3-Long, 4-Double
    If Dt <> 0 Then
        Select Case Dt
            Case 1
                NullCheck = CStr(Trim(NullCheck))
            Case 2
                NullCheck = CInt(Trim(NullCheck))
            Case 3
                NullCheck = CLng(Trim(NullCheck))
            Case 4
                NullCheck = CDbl(Trim(NullCheck))
        End Select
    End If
End Function

Public Function SymbolTxtpart2(MainTextData As Variant, ConditionSymbol As String) As Variant
    '==========================================================================='
    ' Function 疙 : SymbolTxtPart2                                              '
    ' 扁       瓷 : 巩磊凯函荐狼 吝埃 扁龋甫 备盒窍咯 力芭窍哥,                 '
    '               力芭茄 磊丰甫 窃荐疙俊 犁 沥狼 钦聪促.                      '
    ' In Message  : MainTextData As String- 巩磊凯函荐                          '
    '               ConditionSymbol As String - 备盒扁龋                        '
    ' OutMessage  : FnMessage - 箭磊 棺 True,false蔼 馆券 鸥涝 Variant          '                                          '                                                '
    '==========================================================================='

Dim i As Integer
Dim k As Integer
Dim st1 As Integer
Dim strTemp1 As String
Dim strTemp2 As String
Dim strTemp3 As String

strTemp1 = MainTextData
        st1 = 1
        k = 0
        For i = 1 To Len(strTemp1)
            k = k + 1
            If Mid(strTemp1, k, 1) = ConditionSymbol Then
                strTemp1 = Mid(strTemp1, 1, k - 1) + Mid(strTemp1, k + 1, Len(strTemp1) - k)
                k = k - 1
            End If
        Next i
        SymbolTxtpart2 = strTemp1
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_subValueCheck
'   2.Name         : Check Min、Max、Tgt of Master
'   3.Input  Value : Object(dMin、dMax、dTgt)
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 07 .10
'   7.Modify Date  :
'   8.Comment      : Check Min、Max、Tgt
'---------------------------------------------------------------------------------------
Public Function Gf_subValueCheck(ByVal dMin As Object, ByVal dMax As Object, Optional dTgt As Object) As Boolean
    
    Dim min, max, tgt As Integer
   
    If dTgt Is Nothing Then
        If Trim(dMin.Value) = "" And Trim(dMax.Value) = "" Then
            Gf_subValueCheck = True
            Exit Function
        End If
        min = Val(dMin.Value)
        max = Val(dMax.Value)
        If min <= max Then
            Gf_subValueCheck = True
        Else
            Gf_subValueCheck = False
        End If
    
    Else
    
        If Trim(dMin.Value) = "" And Trim(dMax.Value) = "" And Trim(dTgt.Value) = "" Then
            Gf_subValueCheck = True
            Exit Function
        End If
        min = Val(dMin.Value)
        max = Val(dMax.Value)
        tgt = Val(dTgt.Value)
    
        If min <= tgt And max >= tgt Then
            Gf_subValueCheck = True
        Else
            Gf_subValueCheck = False
        End If
    End If
     

    If Gf_subValueCheck = False Then
        Call Gp_MsgBoxDisplay("数据校验错误－请检查包含最小值、最大值或目标值的项目", "I")
        dMin.SetFocus
    End If
                
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_Sp_subValueCheck
'   2.Name         : Check Min、Max、Tgt of Master
'   3.Input  Value : Collection(SC),Long(iRow,iMin,iMax,iTgt),Object(dSetobject),String(sCMsg)
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 07 .10
'   7.Modify Date  :
'   8.Comment      : Check Min、Max、Tgt
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_subValueCheck(ByVal Sc As Collection, ByVal iRow As Long, ByVal iMin As Long, ByVal iMax As Long, Optional sCMsg As String, Optional dSetObject As Object, Optional iTgt As Long = 0) As Boolean
    
    Dim min, max, tgt As Integer
   
    With Sc.Item("Spread")
        If iTgt = 0 Then
            .Row = iRow
            .Col = iMin
            min = Val(.Value)
            .Col = iMax
            max = Val(.Value)
            If min = 0 And max = 0 Then
                Gf_Sp_subValueCheck = True
                Exit Function
            Else
                If min <= max Then
                    Gf_Sp_subValueCheck = True
                    Exit Function
                Else
                    Gf_Sp_subValueCheck = False
                End If
            End If
        Else
            .Row = iRow
            .Col = iMin
            min = Val(.Value)
            .Col = iMax
            max = Val(.Value)
            .Col = iTgt
            tgt = Val(.Value)
            If min <= tgt And max >= tgt Then
                Gf_Sp_subValueCheck = True
                Exit Function
            Else
                Gf_Sp_subValueCheck = False
            End If
            
        End If
    End With
    

    If Gf_Sp_subValueCheck = False Then
        
        If Not (Trim(sCMsg) = "") Then
            Call Gp_MsgBoxDisplay("数据校验错误-请检查" + sCMsg + "的最小值、最大值或目标值数据是否正确!", "I")
        Else
            Call Gp_MsgBoxDisplay("数据校验错误-请检查光标处的最小值、最大值或目标值数据是否正确!", "I")
        End If
                
        Call GP_SELECT_ROW(Sc.Item("Spread"), iRow)
        If Not (dSetObject Is Nothing) Then
            dSetObject.SetFocus
        End If
    End If
                
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_Necessary_Value_Check
'   2.Name         : Check necessary values of the Master whether inputted by users
'   3.Input  Value : Object(dCheckObject,dSetObject),String(sCMsg)
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 08 .02
'   7.Modify Date  :
'   8.Comment      : Check necessary values
'---------------------------------------------------------------------------------------
Public Function GF_Necessary_Value_Check(ByVal dCheckObject As Object, Optional sCMsg As String, Optional dSetObject As Object) As Boolean
        
    GF_Necessary_Value_Check = True
    
    
    If TypeOf dCheckObject Is TextBox Then
        If dCheckObject.Text = "" Or Len(Trim(dCheckObject.Text)) = 0 Then
            GF_Necessary_Value_Check = False
        End If
    ElseIf TypeOf dCheckObject Is sidbEdit Then
        If dCheckObject.Value = 0 Or Len(Trim(dCheckObject)) = 0 Then
            GF_Necessary_Value_Check = False
        End If
    End If

    If GF_Necessary_Value_Check = False Then
        If Not (Trim(sCMsg) = "") Then
            Call Gp_MsgBoxDisplay("数据校验错误-请检查" + sCMsg + "数据是否输入!", "I")
        Else
            Call Gp_MsgBoxDisplay("数据校验错误-请在光标处输入数据!", "I")
        End If
        If Not (dSetObject Is Nothing) Then
            dSetObject.SetFocus
        Else
            dCheckObject.SetFocus
        End If
    End If
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_Sp_Necessary_Value_Check
'   2.Name         : Check necessary values of the Spread whether inputted by users
'   3.Input  Value : Object(Sc,dSetObject),String(sCMsg),Long(iRow,iCol)
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 08 .02
'   7.Modify Date  :
'   8.Comment      : Check necessary values
'---------------------------------------------------------------------------------------
Public Function GF_Sp_Necessary_Value_Check(ByVal Sc As Collection, ByVal iRow As Long, ByVal iCol As Long, Optional sCMsg As String, Optional dSetObject As Object, Optional bMsg As Boolean = True) As Boolean
        
    GF_Sp_Necessary_Value_Check = True
    
    Sc.Item("Spread").Row = iRow
    Sc.Item("Spread").Col = iCol
    
    If Sc.Item("Spread").Text = "" Or Len(Trim(Sc.Item("Spread").Text)) = 0 Or Sc.Item("Spread").Value = 0 Then
        GF_Sp_Necessary_Value_Check = False
    End If

    If GF_Sp_Necessary_Value_Check = False Then
        If bMsg Then
            If Not (Trim(sCMsg) = "") Then
                Call Gp_MsgBoxDisplay("数据校验错误-请检查" + sCMsg + "数据是否输入!", "I")
            Else
                Call Gp_MsgBoxDisplay("数据校验错误-请在光标处输入数据!", "I")
            End If
        End If
        Call GP_SELECT_ROW(Sc.Item("Spread"), iRow)
        If Not (dSetObject Is Nothing) Then
            dSetObject.SetFocus
        End If
    End If
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Goto_Row
'   2.Name         : Select Row of Spread when date posted by user
'   3.Input  Value : Spread(ss),Long(iOldMaxRow,iRow)
'   4.Return Value : No Return
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 08 .03
'   7.Modify Date  :
'   8.Comment      : Select Row
'---------------------------------------------------------------------------------------
Public Sub Gp_Goto_Row(ByVal ss As vaSpread, ByVal iOldMaxRow As Long, ByVal iRow As Long)
            
            If ss.MaxRows < iOldMaxRow Then
                Call GP_SELECT_ROW(ss, ss.MaxRows)
            Else
                Call GP_SELECT_ROW(ss, iRow)
            End If

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_subMasterLock
'   2.Name         : Locked iControl when quality design state is "A" or "a"
'   3.Input  Value : Collection(cMc),String(sState)
'   4.Return Value : No Return
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 08 .05
'   7.Modify Date  :
'   8.Comment      : Locked iControl
'---------------------------------------------------------------------------------------
Public Sub Gf_subMasterLock(ByVal cMc As Collection, ByVal sState As String)
    If sState = "A" Or sState = "a" Then
        Call Gp_Ms_ControlLock(cMc("iControl"), True)
    Else
        Call Gp_Ms_ControlLock(cMc("iControl"), False)
    End If
End Sub


'--------------------------------------------------------------------------------------------------------
'   1.ID           : Gf_Control_text_Up
'   2.Name         : Changing lowercase of controls to capital
'   3.Input  Value : Object(oControl),Long(iRow,iCol)-Optional Parameters When oControl's type is Spread
'   4.Return Value : No Return
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 08 .16
'   7.Modify Date  :
'   8.Comment      : Changing lowercase to capital
'--------------------------------------------------------------------------------------------------------
Public Sub Gf_Control_text_Up(ByVal oControl As Object, Optional iRow As Long = 0, Optional iCol As Long = 0)
    Dim sMyCtrText As String
    
        If TypeOf oControl Is TextBox Then
            sMyCtrText = oControl.Text
            sMyCtrText = UCase(sMyCtrText)
            oControl.Text = sMyCtrText
            oControl.SelStart = Len(Trim(oControl.Text))
        ElseIf TypeOf oControl Is sitxEdit Then
            sMyCtrText = oControl.Text
            sMyCtrText = UCase(sMyCtrText)
            oControl.Text = sMyCtrText
        ElseIf TypeOf oControl Is vaSpread Then
            If iRow = 0 Or iCol = 0 Then
                Exit Sub
            Else
                With oControl
                    .Col = iCol
                    .Row = iRow
                    If .CellType = SS_CELL_TYPE_EDIT Then
                        sMyCtrText = .Text
                        sMyCtrText = UCase(sMyCtrText)
                        .Text = sMyCtrText
                    Else
                        Exit Sub
                    End If
                End With
            End If
        Else
            Exit Sub
        End If
End Sub



'---------------------------------------------------------------------------------------
'   1.ID           : GS_Combo_SS_ADD
'   2.Name         : ADD DATA TO COMBO_Sprade
'   3.Input  Value :
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : ADD DATA TO COMBO_Sprade
'---------------------------------------------------------------------------------------
Public Sub GS_Combo_SS_ADD(ByVal sSQL As String, ByVal vSP As vaSpread)

On Error GoTo Error_Rtn

    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim i_Cyc_Row As Integer
    Dim i_Cyc_Col As Integer
    Dim iHeight As Integer
    
    
    If sSQL = "" Or Trim(sSQL) = "" Then
        Exit Sub
    End If
    
    iHeight = 300
    
    Screen.MousePointer = vbHourglass
    
    Set AdoRs = New adodb.Recordset
    
 
    AdoRs.Open sSQL, M_CN1, adOpenKeyset
    
     With vSP
     
        .MaxRows = 0
    
    If Not AdoRs.EOF Then

'        oForm.cbo_THK_MAX.Clear
'        oForm.cbo_THK_MIN.Clear
        
        ArrayRecords = AdoRs.GetRows
        
        .MaxCols = UBound(ArrayRecords, 1) + 1
        .MaxRows = UBound(ArrayRecords, 2) + 1
     '   .RowHeight = 300
        .Height = (.MaxRows * iHeight) + 15
        
        For i_Cyc_Row = 0 To UBound(ArrayRecords, 2)
        
            .Row = i_Cyc_Row + 1
            For i_Cyc_Col = 0 To UBound(ArrayRecords, 1)
                .Col = i_Cyc_Col + 1
                .Text = Val(ArrayRecords(i_Cyc_Col, i_Cyc_Row))
            Next i_Cyc_Col
            
        Next i_Cyc_Row
        
        
    Else
     
     .MaxRows = 1
                                    
    End If
    
'     .MaxRows = 5
'     .Height = 1500
'     .Row = 1
'     .Col = 1
'     .Text = "11"
'     .Col = 2
'     .Text = "12"
    
    End With
                
    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    
Error_Rtn:

    Screen.MousePointer = vbDefault

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GS_Combo_SS_ADD
'   2.Name         : Set COMBO_Sprade Color
'   3.Input  Value : vSP is a Spread
'   4.Return Value : None
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2005. 05 .13
'   7.Modify Date  :
'   8.Comment      : ADD DATA TO COMBO_Sprade
'---------------------------------------------------------------------------------------

Public Sub GS_ssBackColorSet(ByVal vSP As vaSpread)

    Dim i As Integer

    
    For i = 1 To vSP.MaxRows

        Call Gp_Sp_RowColor(vSP, i, vbBlack, &HC0FFFF)
        
    Next i


End Sub


Public Sub ComboAdd(ByVal cComboBox As Collection, ByVal sSQL As String)

On Error GoTo Error_Rtn

    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i, j As Integer

'    If Trim(txt_NISCO_QUALITY_NO.Text) = "" Then  'Or Trim(txt_APP_DATE.Text) = "" Then
'        Exit Sub
'    End If

    Screen.MousePointer = vbHourglass

    Set AdoRs = New adodb.Recordset
    sQuery = sSQL
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not AdoRs.EOF Then
        
        For i = 1 To cComboBox.COUNT
            cComboBox(i).Clear
        Next i

        ArrayRecords = AdoRs.GetRows

        For i = 0 To UBound(ArrayRecords, 2)
            
            For j = 1 To cComboBox.COUNT
                cComboBox(j).AddItem ArrayRecords(j - 1, i)
            Next j
        
        Next i

    End If

    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault

Error_Rtn:

End Sub

Public Function Val_zero(ByVal f_Val As Single) As String
    
    If f_Val < 1 Then
        Val_zero = Trim("0" + str(f_Val))
    Else
        Val_zero = Trim(str(f_Val))
    End If
    
End Function

Public Function Find_SMP_LOC(ByVal s_SMP_NO As String) As String
    Dim sQuery      As String
    Dim sErrCode    As String
    Dim sProd_Table As String
    Dim AdoRs       As adodb.Recordset
    
    sErrCode = 0
    Set AdoRs = New adodb.Recordset
'-------------------------------------------------------------------------------------------
    sQuery = "SELECT DISTINCT SMP_CUT_LOC FROM QP_TEST_HEAD WHERE SMP_NO = '" + Trim(s_SMP_NO) + "'"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        sErrCode = 1
    Else
        Find_SMP_LOC = AdoRs.Fields(0).Value
        
        If Find_SMP_LOC = "Y" Then
           Find_SMP_LOC = ""
           Call MsgBox("该试样号需要头尾取样，请在录入取样位置（T：头部，B：尾部）！", vbOKOnly, "系统提示")
        End If
        
        sErrCode = 0
    End If
'-------------------------------------------------------------------------------------------
    If sErrCode = 0 Then
        AdoRs.Close
        Set AdoRs = Nothing
        Exit Function
    Else
        AdoRs.Close
        Set AdoRs = Nothing
        Call MsgBox("试样号错误 - 试样号不存在，请检查！", vbOKOnly, "系统提示")
        Exit Function
    End If
End Function

Public Function CERT_PONO_SET(ByVal s_ISP_SHP_NO As String) As String
    
    Dim s_MyPono As String
    
    
    Select Case Trim(s_ISP_SHP_NO)
        Case "GO060706010", "GO060706011", "GO060706012", "GO060706013", _
             "GO060706014", "GO060706015", "GO060706016", "GO060706017"
            
            s_MyPono = "06JTE/P009LOT1"
            
        Case "GO060619004", "GO060619005", "GO060619006", "GO060619007"
            
            s_MyPono = "06JTE/P009LOT2A"
        
        Case "GO060619008", "GO060619009", "GO060619010", "GO060619011", "GO060801003", _
             "GO060801004", "GO060808002"
        
            s_MyPono = "06JTE/P009LOT2B"
        
        Case "GO060619012", "GO060619013", "GO060619014", "GO060619015", "GO0619016", _
             "GO060801005", "GO060801006", "GO060801007", "GO060619016" '--hjd
        
            s_MyPono = "06JTE/P009LOT3A"
        
        Case "GO060619017", "GO060619018", "GO060619019", "GP060619020", _
             "GO060801008", "GO060801009", "GO060619020" '--hjd
        
            s_MyPono = "06JTE/P009LOT3B"
        
        Case "GO060706018", "GO060706019", "GO060706020", "GO060706021", "GO060706022", "GO060706023", _
             "GO060706024", "GO060801010", "GO060801011"
        
            s_MyPono = "06JTE/P009LOT4"
        
        Case "GO060706025", "GO060706026", "GO060706027"
    
            s_MyPono = "06JTE/P009LOT5"
        
        Case "GO060707001", "GO060707002", "GO060707003", "GO060707004", "GO060707005", "GO060707006", _
             "GO060707007", "GO060707008", "GO060707009", "GO060707010", "GO060707011", "GO060707012", _
             "GO060707013", "GO060707014", "GO060707015", "GO060707016", "GO060801012"
            
            s_MyPono = "06NGE/P057LOT1"
             
        Case "GO060707017", "GO060707018", "GO060707019", "GO060707020", "GO060707021", "GO060707022", _
             "GO060707023", "GO060707024", "GO060707025", "GO060707026", "GO060707027", "GO060707028", _
             "GO060707029", "GO060707030", "GO060707031", "GO060801013", "GO060711002", "GO060711003"
             
            s_MyPono = "06NGE/P057LOT2"
             
        Case "GO060707032", "GO060707033", "GO060707034", "GO060707035", "GO060707036", "GO060707037", _
             "GO060707038", "GO060707039", "GO060707040", "GO060707041", "GO060707042", "GO060707043", _
             "GO060707044", "GO060707045", "GO060707046", "GO060707047", "GO060707048", "GO060801014", _
             "GO060801015", "GO060801016", "GO060801017", "GO060801018", "GO060801019", "GO060801020", _
             "GO060801021"
            
            s_MyPono = "06NGE/P057LOT3"
             
        Case "GO060707049", "GO060707050", "GO060707051", "GO060707052", "GO060707053", "GO060707054", _
             "GO060707055", "GO060707056", "GO060707057", "GO060707058", "GO060707059", "GO060707060", _
             "GO060707061", "GO060707062", "GO060707063", "GO060707064", "GO060707065", "GO060707066", _
             "GO060707067", "GO060707068", "GO060707069", "GO060707070", "GO060707071", "GO060707072", _
             "GO060707073", "GO060707074", "GO060707075", "GO060707076", "GO060707077", "GO060707078", _
             "GO060707079", "GO060707080", "GO060801022", "GO060801023", "GO060801024"
        
            s_MyPono = "06NGE/P057LOT4"
            
        Case "GO060804002", "GO060804003", "GO060804004", "GO060804005", "GO060804006"
        
            s_MyPono = "06NGE/P065LOT1A"
        
        Case "GO060801033", "GO060801034", "GO060801035"
        
            s_MyPono = "06NGE/P065LOT1B"
        
        Case "GO060804007", "GO060804008", "GO060804009", "GO060804010", "GO060804011", "GO060804012", _
             "GO060804013", "GO060804014", "GO060804015", "GO060804016"
        
            s_MyPono = "06NGE/P065LOT1C"
        
        Case "GO060801036", "GO060801037", "GO060801038", "GO060801039", "GO060801040", "GO060801041", _
             "GO060801042", "GO060801043", "GO060801044", "GO060801045", "GO060801046", "GO060801047", _
             "GO060801048", "GO060801049", "GO060801050"
        
            s_MyPono = "06NGE/P065LOT1D"
        
        Case "GO060801051", "GO060801052", "GO060801053", "GO060801054", "GO060801055", "GO060801056", _
             "GO060801057", "GO060801058", "GO060801059", "GO060801060", "GO060801061", "GO060801062", _
             "GO060801063", "GO060801064", "GO060801065", "GO060801066", "GO060801067", "GO060801068", _
             "GO060801069", "GO060801070", "GO060801071", "GO060801072", "GO060801073", "GO060801074", _
             "GO060801075", "GO060801076"
        
            s_MyPono = "06NGE/P065LOT2"
        
        Case "GO060801077", "GO060801078", "GO060801079", "GO060801080", "GO060801081", "GO060801082", _
             "GO060801083", "GO060801084", "GO060801085", "GO060801086", "GO060801087", "GO060801088"
             
            s_MyPono = "06NGE/P065LOT3"
        
        Case "GO060801089", "GO060801090", "GO060801091", "GO060801092", "GO060801093", "GO060801094", _
             "GO060801095", "GO060801096", "GO060801097", "GO060801098", "GO060801099", "GO060801100", _
             "GO060801101", "GO060801102", "GO060801103", "GO060801104", "GO060801105", "GO060801106", _
             "GO060801107", "GO060801108", "GO060801109", "GO060801110", "GO060801111", "GO060801112", _
             "GO060808001"
            
            s_MyPono = "06NGE/P065LOT4"
            
        Case "GO060823007", "GO060823062", "GO060823063", "GO060823064", "GO060823065", "GO060823066", _
             "GO060823067", "GO060823068", "GO060823069", "GO060823070", "GO060823071", "GO060823072", _
             "GO060823073", "GO060823074", "GO060823075", "GO060823076", "GO060823077", "GO060823078", _
             "GO060823079", "GO060823080", "GO060826001", "GO060826002", "GO060826003", "GO060826004", _
             "GO060826005", "GO060826006", "GO060826007"
             
            s_MyPono = "06NGE/P070LOT1"
            
        Case "GO060802001", "GO060802002", "GO060802003", "GO060802004", "GO060802005", "GO060802006", _
             "GO060802007", "GO060802008", "GO060802009", "GO060802010", "GO060802011", "GO060802012", _
             "GO060802013", "GO060802014", "GO060802015", "GO060802016", "GO060802017", "GO060802018", _
             "GO060802019", "GO060802020", "GO060802021", "GO060802022", "GO060802023", "GO060802024", _
             "GO060802025", "GO060802026", "GO060808003", "GO060804001"
             
            s_MyPono = "06NGE/P070LOT2A"
        
        Case "GO060823036", "GO060823037", "GO060823038", "GO060823039", "GO060823040", "GO060823041", _
             "GO060823042", "GO060823043", "GO060823044", "GO060823045", "GO060823046", "GO060823047", _
             "GO060823048", "GO060823049", "GO060823050", "GO060823051", "GO060823052", "GO060823053", _
             "GO060823054", "GO060823055", "GO060823056", "GO060823057", "GO060823058", "GO060823059", _
             "GO060823060", "GO060823061", "GO060826008", "GO060826009", "GO060826010"
        
            s_MyPono = "06NGE/P070LOT2B"
        
        Case "GO060823009", "GO060823010", "GO060823011", "GO060823012", "GO060823013", "GO060823014", _
             "GO060823015", "GO060823016", "GO060823017", "GO060823018", "GO060823019", "GO060823020", _
             "GO060823021", "GO060823022", "GO060823023", "GO060823024", "GO060823025", "GO060823026", _
             "GO060823027", "GO060823028", "GO060823029", "GO060823030", "GO060823031", "GO060823032", _
             "GO060823033", "GO060823034", "GO060823035", "GO060823134", "GO060823135", "GO060826011", _
             "GO060826012", "GO060826013", "GO060826014", "GO060826015", "GO060826016", "GO060826017", _
             "GO060826018", "GO060826019"
        
            s_MyPono = "06NGE/P070LOT3A"
        
        Case "GO060823105", "GO060823106", "GO060823107", "GO060823108", "GO060823109", "GO060823110", _
             "GO060823111", "GO060823112", "GO060823113", "GO060823114", "GO060823115", "GO060823116", _
             "GO060823117", "GO060823118", "GO060823119", "GO060823120", "GO060823121", "GO060823122", _
             "GO060823123", "GO060823124", "GO060823125", "GO060823126", "GO060823127", "GO060823128", _
             "GO060823129", "GO060823130", "GO060823131", "GO060823132", "GO060823133", "GO060826020", _
             "GO060826021", "GO060826022", "GO060826023", "GO060826024", "GO060826025", "GO060826026", _
             "GO060826027", "GO060826028"
                
            s_MyPono = "06NGE/P070LOT3B"
        
        Case "GO060823081", "GO060823082", "GO060823083", "GO060823084", "GO060823085", "GO060823086", _
             "GO060823087", "GO060823088", "GO060823089", "GO060823090", "GO060823091", "GO060823092", _
             "GO060823093", "GO060823094", "GO060823095", "GO060823096", "GO060823097", "GO060823098", _
             "GO060823099", "GO060823100", "GO060823101", "GO060823102", "GO060823103", "GO060823104", _
             "GO060826029", "GO060826030", "GO060826031", "GO060826032", "GO060826033", "GO060826034", _
             "GO060826035", "GO060826036"
        
            s_MyPono = "06NGE/P070LOT4"
        

        Case Else
    
            s_MyPono = "NO"
    End Select

    CERT_PONO_SET = s_MyPono
    
End Function
'---------------------------------------------------------------------------------------
'   1.ID           : Ship_Input_AUTH
'   2.Name         : Get User's authority for input ship prod
'   3.Input  Value : sSMP_NO , sEMP_ID
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 11 .21
'   7.Modify Date  :
'   8.Comment      :
'---------------------------------------------------------------------------------------
Public Function Ship_Input_AUTH(ByVal sSMP_NO As String, ByVal sEMP_ID As String, ByVal sOldAuthority) As String
    Dim sQuery          As String
    Dim AdoRs           As adodb.Recordset
    Dim sEND_CD         As String
    Dim sPROD_CD        As String
 
 On Error GoTo Error_Rtn

    Set AdoRs = New adodb.Recordset
  
    If Trim(sSMP_NO) = "" Or Len(Trim(sSMP_NO)) = 0 Then
        GoTo Error_Rtn
    End If
' No.1
    sQuery = "SELECT ENDUSE_CD,PROD_CD FROM QP_TEST_HEAD WHERE SMP_NO = '" + sSMP_NO + "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        sEND_CD = AdoRs.Fields(0).Value
        sPROD_CD = AdoRs.Fields(1).Value
    Else
        GoTo Error_Rtn
    End If
    
    AdoRs.Close
'No.2
    If sPROD_CD = "PP" Then
        If sEND_CD <> "F01" Then
            Ship_Input_AUTH = sOldAuthority
            Set AdoRs = Nothing
            Exit Function
        End If
    ElseIf sPROD_CD = "HC" Then
        If sEND_CD <> "F01" Then
            Ship_Input_AUTH = sOldAuthority
            Set AdoRs = Nothing
            Exit Function
        End If
    Else
        Ship_Input_AUTH = sOldAuthority
        Set AdoRs = Nothing
        Exit Function
    End If
'No.3
    sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD_MANA_NO = 'Q0069' AND CD = '" + sEMP_ID + "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        Ship_Input_AUTH = "1111"
    Else
        Ship_Input_AUTH = "1000"
    End If

    AdoRs.Close

    Set AdoRs = Nothing

    Exit Function

Error_Rtn:
    
    Ship_Input_AUTH = sOldAuthority
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
    
End Function
''---------------------------------------------------------------------------------------
''   1.ID           : Chem_Check_AUTH
''   2.Name         : Get User's authority for input ship prod
''   3.Input  Value : sEMP_ID
''   4.Return Value : String
''   5.Writer       : sun bin
''   6.Create Date  : 2008. 08 .25
''   7.Modify Date  :
''   8.Comment      :
''---------------------------------------------------------------------------------------
Public Function Chem_Check_AUTH(ByVal sEMP_ID As String, ByVal sOldAuthority) As String
    Dim sQuery          As String
    Dim AdoRs           As adodb.Recordset
    Dim sEND_CD         As String

 On Error GoTo Error_Rtn

    Set AdoRs = New adodb.Recordset

    sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD_MANA_NO = 'Q0077' AND CD = '" + sEMP_ID + "'"

    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not (AdoRs.BOF And AdoRs.EOF) Then
        Chem_Check_AUTH = "1111"
    Else
        Chem_Check_AUTH = "1000"
    End If

    AdoRs.Close

    Set AdoRs = Nothing

    Exit Function

Error_Rtn:

    Chem_Check_AUTH = sOldAuthority
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Expo_SMP_CHECK
'   2.Name         : Get sampling is Expo sampling or no Expo sampling
'   3.Input  Value : sSMP_NO
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 12 .03
'   7.Modify Date  :
'   8.Comment      :
'---------------------------------------------------------------------------------------
Function Expo_SMP_Check(ByVal sSMP_NO As String) As Boolean
    Dim sPROD_CD        As String
    Dim AdoRs           As adodb.Recordset
    Dim sQuery          As String
    Dim sLast_SMP_NO    As String
 
    On Error GoTo Error_Rtn
    
    Set AdoRs = New adodb.Recordset
    
    sLast_SMP_NO = Right(Trim(sSMP_NO), 2)
    
    sQuery = "Select Prod_CD From QP_TEST_HEAD Where SMP_NO = '" + sSMP_NO + "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        sPROD_CD = AdoRs.Fields(0).Value
    Else
        GoTo Error_Rtn
    End If
    
    AdoRs.Close
    
    If UCase(sPROD_CD) = "PP" Then
        If sLast_SMP_NO = "00" Then
            Expo_SMP_Check = True
        Else
            Expo_SMP_Check = False
        End If
    ElseIf UCase(sPROD_CD) = "HC" Then
        If sLast_SMP_NO = "99" Then
            Expo_SMP_Check = True
        Else
            Expo_SMP_Check = False
        End If
    Else
        GoTo Error_Rtn
    End If
        
    Set AdoRs = Nothing
    
    Exit Function
    
Error_Rtn:
    Expo_SMP_Check = False
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : SMP_PROD_CHECK
'   2.Name         : Get sampling'prod_cd value
'   3.Input  Value : sSMP_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007.02.14
'   7.Modify Date  :
'   8.Comment      :
'---------------------------------------------------------------------------------------
Function SMP_PROD_Check(ByVal sSMP_NO As String) As String
    Dim sPROD_CD        As String
    Dim AdoRs           As adodb.Recordset
    Dim sQuery          As String
 
    On Error GoTo Error_Rtn
    
    Set AdoRs = New adodb.Recordset
      
    sQuery = "Select Prod_CD From QP_TEST_HEAD Where SMP_NO = '" + sSMP_NO + "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        sPROD_CD = AdoRs.Fields(0).Value
    Else
        GoTo Error_Rtn
    End If
    
    SMP_PROD_Check = sPROD_CD
    
    AdoRs.Close
    
    Exit Function
Error_Rtn:
    SMP_PROD_Check = "ER"
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : Gf_PLT_Authority
'   2.Name         : Program ID Authority Check
'   3.Input  Value : sPLTID String
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .10
'   7.Modify Date  :
'   8.Comment      : PLANT Authority Check
'---------------------------------------------------------------------------------------
Public Function Gf_PLT_Authority(sPgmId As String) As String
    
On Error GoTo Pgm_Authority_Error

    Dim sQuery As String
     
    If sUserID = "1JS6001" Or sUserID = "1JS6002" Or sUserID = "1JS6003" Or sUserID = "1JS6005" Then
        Gf_PLT_Authority = "**"
    Else
               
        'Authority Check
        sQuery = "SELECT PLT FROM ZP_EMPLOYEE WHERE EMP_ID = '" + sUserID + "' "
        Gf_PLT_Authority = Gf_CodeFind(M_CN1, sQuery)
        
    End If
    
    Exit Function

Pgm_Authority_Error:

    Gf_PLT_Authority = "FAIL"

End Function
