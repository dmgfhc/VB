Attribute VB_Name = "basB"
Option Explicit

Public Const SS_TEXTTIP_FIXEDFOCUSONLY = 3
Public Const SS_TEXTTIP_FLOATINGFOCUSONLY = 4

Global sOrderNo As String                    '������ OrderNo
Global sOrderItem As String                  '�������к� OrderItem
Global sSampCd As String                     'ȡ������
Global sSampSearch As String                 'ȡ������ Search Text

Global arrSampCd1() As Variant               'ȡ������1
Global arrSampCd2() As Variant               'ȡ������2
Global arrSampCd3() As Variant               'ȡ������3
Global arrSampCd4() As Variant               'ȡ������4
Global arrSampCd5() As Variant               'ȡ������5

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
Public Function GF_GetCellMaxLength(ss1 As vaSpread, iRow As Long, iCol As Long) As Double
    With ss1
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
Public Function Gf_Get_Cell_Value(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With ss1
        .Row = iRow
        .Col = iCol
        Gf_Get_Cell_Value = .Value
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
Public Function GF_GET_CELL_VALUE2(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With ss1
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
Public Function Gf_GetCellText(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With ss1
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
Public Function Gf_GetCellNullCheck(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    With ss1
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
Public Function Gf_GetCellText2(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    Dim str As String
    With ss1
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
Public Sub Gp_SetRowColor(ss1 As vaSpread, ByVal iRow As Long)

    Dim i As Long

    Call Gp_Sp_EvenRowBackcolor(ss1)

    With ss1
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
Public Sub Gp_SetCellFormula(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long, ByVal sFormula As String)
    With ss1
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
Public Sub GP_SET_CELL_VALUE(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long, sText As Variant)
    
    If iRow <= 0 Then Exit Sub
    
    With ss1
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
Public Sub GP_SET_CELL_VALUE2(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long, ByVal iCnt As Integer, sText1 As Variant, Optional ByVal sText2 As Variant, Optional ByVal sText3 As Variant, Optional ByVal sText4 As Variant, Optional ByVal sText5 As Variant, Optional ByVal sText6 As Variant, Optional ByVal sText7 As Variant, Optional ByVal sText8 As Variant, Optional ByVal sText9 As Variant)
    
    Dim str As String
    Dim TXT_CNT As Integer
    
    If iRow <= 0 Then Exit Sub
    
    With ss1
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
Public Sub GP_ChemCode_RowHeader_Clear(ByVal sFormName As String, ss1 As vaSpread, ByVal iRow As Integer, ByVal iCol As Integer)
    
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
       
       With ss1
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
            .Text = "��"
            '.Text = ">>"
        End If
        
         
       End With
       
       'ss1.SetFocus
         
'       Call GP_SetRowHeaderClear(ss1, iRow, iColNo)
       
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
Public Sub GP_SetRowHeaderClear(ss1 As vaSpread, ByVal iRow As Long, Optional ByVal iCol As Integer = 0)
    
    Dim i As Long
    
    With ss1
    
        For i = 1 To .MaxRows
                
            .Row = i
            .Col = iCol
            If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
                .Text = ""
            End If
        Next i
        
        .Row = iRow: .Col = iCol
        
        If .Text <> "Input" And .Text <> "Update" And .Text <> "Delete" Then
            .Text = "��"  '�����������
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
    
    Dim iCount As Integer

    For iCount = 1 To iControl.Count
        
        iControl.Item(iCount).BackColor = vbWhite
        
    Next iCount
    
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
        
            Case "STDSPEC"              '��׼��
                Call Gf_StdSPEC_DD(M_CN1, KeyCode)
            
            Case "CUST_CD"              '�ͻ�
                DD.nameType = "1"
                Call Gf_Customer_DD(M_CN1, KeyCode)
            
            Case "ENDUSE_CD"            '������;
                Call Gf_Usage_DD(M_CN1, KeyCode)
            
            Case "STLGRD"               '����
                Call Gf_Stlgrd_DD(M_CN1, KeyCode)
            
            Case "CUST_SPEC_NO"         '�ͻ�����Ҫ����
                Call Gf_Cust_STD_DD(M_CN1, KeyCode)
            
            Case "NISCO_QUALITY_NO"     '�����ʱ��
                Call Gf_Nisco_STD_DD(M_CN1, KeyCode)
                
            Case "MLT_STD_NO"           '���ֹ�̱��
                Call Gf_Melt_STD_DD(M_CN1, KeyCode)
                
            Case "MILL_STD_NO"          '���ֹ�̱��
                Call Gf_Roll_STD_DD(M_CN1, KeyCode)
            
            Case "DEV_STD_CD"          '�����Խ���������׼
                Call Gf_STD_DELV_DD(M_CN1, KeyCode)
                            
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
        
            Case "STDSPEC"      '��׼��
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
Public Function Gf_Procedure_Exec(sQuery As String) As Boolean

On Error GoTo ExecQuery_ERROR

    Dim ret_Result_ErrCode As String
    Dim ret_Result_ErrMsg As String
    Dim adoCmd As ADODB.Command
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
    Set adoCmd = New ADODB.Command
    
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
        
        sErrMessg = "���� ���� : " & ret_Result_ErrCode & vbCrLf & "���� ��Ϣ : " & ret_Result_ErrMsg
        
        Call Gp_MsgBoxDisplay(sErrMessg)
        
        Set adoCmd = Nothing
        Gf_Procedure_Exec = False
    
        Exit Function
    Else
        ret_Result_ErrMsg = NullCheck(adoCmd("arg_e_msg"), "")
        sErrMessg = ret_Result_ErrMsg
        
       ' Call Gp_MsgBoxDisplay(sErrMessg)
        
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

    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i As Integer
    Dim sTable As String
    
    If Trim(oForm.txt_stdspec.Text) = "" Or Trim(oForm.txt_stdspec_yy.Text) = "" Then
        Exit Sub
    End If

    oForm.cbo_THK_MAX.Clear
    oForm.cbo_THK_MIN.Clear
    
    Screen.MousePointer = vbHourglass
    
    Set adoRs = New ADODB.Recordset
    
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
   
    sQuery = "SELECT  DISTINCT THK_MIN , THK_MAX FROM " + sTable + " WHERE STDSPEC = '" + Trim(oForm.txt_stdspec.Text) + "' AND STDSPEC_YY = '"
    sQuery = sQuery + Trim(oForm.txt_stdspec_yy.Text) + "'"
 
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not adoRs.EOF Then

'        oForm.cbo_THK_MAX.Clear
'        oForm.cbo_THK_MIN.Clear
        
        ArrayRecords = adoRs.GetRows
        
        For i = 0 To UBound(ArrayRecords, 2)
            
            oForm.cbo_THK_MIN.AddItem ArrayRecords(0, i)
            oForm.cbo_THK_MAX.AddItem ArrayRecords(1, i)
        Next i
                                    
    End If
                
    adoRs.Close
    Set adoRs = Nothing
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

    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i As Integer
    Dim sTable As String
    Dim iHeight As Integer
    
    If Trim(oForm.txt_stdspec.Text) = "" Or Trim(oForm.txt_stdspec_yy.Text) = "" Then
        Exit Sub
    End If
    
    iHeight = 245
    
    Screen.MousePointer = vbHourglass
    
    Set adoRs = New ADODB.Recordset
    
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
   
    sQuery = "SELECT  DISTINCT THK_MIN , THK_MAX FROM " + sTable + " WHERE STDSPEC = '" + Trim(oForm.txt_stdspec.Text) + "' AND STDSPEC_YY = '"
    sQuery = sQuery + Trim(oForm.txt_stdspec_yy.Text) + "'"
 
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    
     With oForm.ss2
     
        .MaxRows = 0
    
    If Not adoRs.EOF Then

'        oForm.cbo_THK_MAX.Clear
'        oForm.cbo_THK_MIN.Clear
        
        ArrayRecords = adoRs.GetRows
                           
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
                
    adoRs.Close
    Set adoRs = Nothing
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

    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i As Integer
    Dim sTable As String
        
    Screen.MousePointer = vbHourglass
    
    Set adoRs = New ADODB.Recordset
    
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
   
    sQuery = "SELECT  DISTINCT THK_MIN , THK_MAX FROM " + sTable + " WHERE STDSPEC = '" + Trim(oForm.txt_stdspec.Text) + "' AND STDSPEC_YY = '"
    sQuery = sQuery + Trim(oForm.txt_stdspec_yy.Text) + "'"
 
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not adoRs.EOF Then

        ArrayRecords = adoRs.GetRows
                                                                   
    End If
    
    Gf_Thick_Mix_Max = ArrayRecords
    
'    End With
                
    adoRs.Close
    Set adoRs = Nothing
    Screen.MousePointer = vbDefault
    
Error_Rtn:
    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GP_SET_THK_MIN_MAX_VALUE
'   2.Name         : ����� MIN , MAX VALUE
'   3.Input  Value : MIN , MAX , TAEGET
'   4.Return Value : None
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .30
'   7.Modify Date  :
'   8.Comment      : ����� MIN , MAX VALUE
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
'   2.Name         : ����ֵ , ����ֵ , Ŀ��ֵ Check
'   3.Input  Value : MIN , MAX , TAEGET
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : ����ֵ , ����ֵ , Ŀ��ֵ Check
'---------------------------------------------------------------------------------------
Public Function GF_MIN_MAX_TARGET_CHECK(ByVal dMin As Object, ByVal dMax As Object, ByVal dTgt As Object) As Boolean

    If Trim(dMin.Text) = "" And Trim(dMax.Text) = "" And Trim(dTgt.Text) Then
        GF_MIN_MAX_TARGET_CHECK = True
        Exit Function
    End If

    If Trim(dMin.Text) <= Trim(dTgt.Text) And Trim(dMax.Text) >= Trim(dTgt.Text) Then
        GF_MIN_MAX_TARGET_CHECK = True
    Else
        GF_MIN_MAX_TARGET_CHECK = False
        Call Gp_MsgBoxDisplay("�������ݴ���,������Χ!", "I")
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
Public Function GF_CHEM_SEQ(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
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

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("�����ֵ�������Ч.....", "I")
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
            
            If DD.rControl.Count > 1 Then
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
    Dim adoRs As ADODB.Recordset
        
    Screen.MousePointer = vbHourglass
                
    sQuery = "SELECT CHEM_COMP_SEQ,CHEM_COMP_CD,CHEM_LEN FROM QP_CHEM_SEQ ORDER BY CHEM_COMP_SEQ"
    
    Set adoRs = New ADODB.Recordset
    
    adoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If adoRs.BOF Or adoRs.EOF Then
        adoRs.Close
        Set adoRs = Nothing
        Screen.MousePointer = 0
        GF_GetChemicalCode = Null
        Exit Function
    End If
        
    GF_GetChemicalCode = adoRs.GetRows
        
    adoRs.Close
    Set adoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

Error_Rtn:

    Set adoRs = Nothing
    GF_GetChemicalCode = Null

    Screen.MousePointer = vbDefault
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_GetSampleCode
'   2.Name         : ȡ���������
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : ȡ���������
'---------------------------------------------------------------------------------------
Public Sub Gp_GetSampleCode()

'On Error GoTo Error_Rtn

    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    
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
            
    arrSampCd1(0, 1) = "ȡ��1��"
    arrSampCd1(1, 1) = "ȡ��2��"
    arrSampCd1(2, 1) = "ȡ��3��"
    arrSampCd1(3, 1) = "ȡ��4��"
    arrSampCd1(4, 1) = "ȡ��5��"
    arrSampCd1(5, 1) = "ȡ��6��"
    arrSampCd1(6, 1) = "ȡ��7��"
    arrSampCd1(7, 1) = "ȡ��8��"
    arrSampCd1(8, 1) = "ȡ��9��"
                
    Set adoRs = New ADODB.Recordset
    
'���ȷ���λ
    sQuery = "SELECT CD , CD_NAME FROM ZP_CD WHERE CD_MANA_NO  = 'Q0021'"
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not adoRs.BOF Then
        arrSampCd2 = adoRs.GetRows
        adoRs.Close
    End If
        
'��ȷ���λ
    sQuery = "SELECT CD , CD_NAME FROM ZP_CD WHERE CD_MANA_NO  = 'Q0022'"
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not adoRs.BOF Then
        arrSampCd3 = adoRs.GetRows
        adoRs.Close
    End If
        

'��ȷ���λ
    sQuery = "SELECT CD , CD_NAME FROM ZP_CD WHERE CD_MANA_NO  = 'Q0023'"
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not adoRs.BOF Then
        arrSampCd4 = adoRs.GetRows
        adoRs.Close
    End If


'�����ߴ����
    sQuery = "SELECT SMP_SIZE_CD , SMP_SPEC FROM QP_SAMP_STD ORDER BY SMP_SIZE_CD"
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not adoRs.BOF Then
        arrSampCd5 = adoRs.GetRows
        'AdoRs.Close
    End If
        
    adoRs.Close
    Set adoRs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Error_Rtn:

    Set adoRs = Nothing

    Screen.MousePointer = vbDefault
    
End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : Gp_SetSampleCode
'   2.Name         : ȡ������ Setting
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : ȡ������ Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_SetSampleCode(oForm As Form)

'On Error GoTo Error_Rtn

    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    
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
Public Sub GF_GetCeqValue(ByVal ss1 As vaSpread, ByVal sFormName As String, ByVal sKnd As String)
    
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
    
    With ss1
    
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
            Case Else
            
                Exit Sub
                
        End Select
            
        
'-------------------- Min value -----------------------------------------------------------
            
            CHEM_C(0) = subGetChemValue(ss1, "C", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_SI(0) = subGetChemValue(ss1, "Si", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_MN(0) = subGetChemValue(ss1, "Mn", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_P(0) = subGetChemValue(ss1, "P", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_S(0) = subGetChemValue(ss1, "S", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_CR(0) = subGetChemValue(ss1, "Cr", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_V(0) = subGetChemValue(ss1, "V", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_MO(0) = subGetChemValue(ss1, "Mo", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_CU(0) = subGetChemValue(ss1, "Cu", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_NI(0) = subGetChemValue(ss1, "Ni", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            CHEM_B(0) = subGetChemValue(ss1, "B", iCode1, iCode2, iCode3, iMin1, iMin2, iMin3)
            
'-------------------- Max value -----------------------------------------------------------
            
            CHEM_C(1) = subGetChemValue(ss1, "C", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_SI(1) = subGetChemValue(ss1, "Si", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_MN(1) = subGetChemValue(ss1, "Mn", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_P(1) = subGetChemValue(ss1, "P", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_S(1) = subGetChemValue(ss1, "S", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_CR(1) = subGetChemValue(ss1, "Cr", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_V(1) = subGetChemValue(ss1, "V", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_MO(1) = subGetChemValue(ss1, "Mo", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_CU(1) = subGetChemValue(ss1, "Cu", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_NI(1) = subGetChemValue(ss1, "Ni", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)
            CHEM_B(1) = subGetChemValue(ss1, "B", iCode1, iCode2, iCode3, iMax1, iMax2, iMax3)

'-------------------- Target value -----------------------------------------------------------
            
        If sKnd = "3" Then
        
            CHEM_C(2) = subGetChemValue(ss1, "C", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_SI(2) = subGetChemValue(ss1, "Si", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_MN(2) = subGetChemValue(ss1, "Mn", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_P(2) = subGetChemValue(ss1, "P", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_S(2) = subGetChemValue(ss1, "S", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_CR(2) = subGetChemValue(ss1, "Cr", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_V(2) = subGetChemValue(ss1, "V", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_MO(2) = subGetChemValue(ss1, "Mo", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_CU(2) = subGetChemValue(ss1, "Cu", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_NI(2) = subGetChemValue(ss1, "Ni", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
            CHEM_B(2) = subGetChemValue(ss1, "B", iCode1, iCode2, iCode3, iTgt1, iTgt2, iTgt3)
        
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
Private Function subGetChemValue(ByVal ss1 As vaSpread, chem As String, ByVal iCode1 As Integer, ByVal iCode2 As Integer, ByVal iCode3 As Integer, ByVal iCol1 As Integer, ByVal iCol2 As Integer, ByVal iCol3 As Integer) As Double
    
    Dim i As Integer
    
    With ss1
        
        For i = 1 To .MaxRows
                            
            If Gf_Get_Cell_Value(ss1, i, iCode1) = chem Then
                subGetChemValue = Val(Gf_Get_Cell_Value(ss1, i, iCol1))
                Exit Function
            End If
        
        Next i
    
        
        For i = 1 To .MaxRows
                    
            If Gf_Get_Cell_Value(ss1, i, iCode2) = chem Then
                subGetChemValue = Val(Gf_Get_Cell_Value(ss1, i, iCol2))
                Exit Function
            End If
        
        Next i
        
        
        For i = 1 To .MaxRows
            
            If Gf_Get_Cell_Value(ss1, i, iCode3) = chem Then
                subGetChemValue = Val(Gf_Get_Cell_Value(ss1, i, iCol3))
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
Public Sub GS_SetChemicalLength(ByVal ss1 As vaSpread, ByVal ArrayRecords As Variant, ByVal sCol As String, ByVal sKnd As String)

    Dim i As Integer
    Dim J As Integer
    Dim dblChem As Double
    Dim iCol(3) As Integer
    Dim iLoop As Integer
    
    
    iCol(1) = CInt(Mid(sCol, 1, 2))
    iCol(2) = CInt(Mid(sCol, 3, 2))
    iCol(3) = CInt(Mid(sCol, 5, 2))
    
    With ss1
        
        For iLoop = 1 To 3
                        
                For i = 1 To .MaxRows
                   
                   .Col = iCol(iLoop): .Row = i
                    
                    For J = 0 To UBound(ArrayRecords, 2)
                        If ArrayRecords(1, J) = Trim(.Text) Then
                            dblChem = ArrayRecords(2, J)
                            Exit For
                        End If
                    Next J
                
                    Call subSetChemLength(ss1, dblChem, .Col, .Row, sKnd)
                
                Next i
        
        Next iLoop
                
    End With

End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : GS_CeqColumnLock
'   2.Name         : Ceq Column -> Lock
'   3.Input  Value : Spread , ���� Code
'   4.Return Value :
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .24
'   7.Modify Date  :
'   8.Comment      : Ceq Column -> Lock
'---------------------------------------------------------------------------------------
Public Sub GS_CeqColumnLock(ByVal ss1 As vaSpread, ByVal sKnd As String, ByVal sCol As String)

    Dim i As Integer
    Dim iCol(3) As Integer
    
    iCol(1) = CInt(Mid(sCol, 1, 2))
    iCol(2) = CInt(Mid(sCol, 3, 2))
    iCol(3) = CInt(Mid(sCol, 5, 2))

    With ss1
    
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
Public Sub GS_SetChemicalSpreadLineColor(ByVal ss1 As vaSpread, ByVal sCol As String)
    
    Dim i As Integer
    Dim iCol(2) As Integer
    
    iCol(1) = CInt(Mid(sCol, 1, 2))
    iCol(2) = CInt(Mid(sCol, 3, 2))
    
    With ss1
        
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
Private Sub subSetChemLength(ByVal ss1 As vaSpread, ByVal dblChem As Double, ByVal Col As Long, ByVal Row As Long, ByVal sKnd As String)

    With ss1
        
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
'   2.Name         : ���� COMMON Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� COMMON Input Check
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
'   2.Name         : ���� MIN , MAX Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� MIN , MAX Input Check
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
'   2.Name         : ���� - ������� Input Check
'   3.Input  Value : IMPACT_SMP_CD,IMPACT_KND,IMPACT_DIR,IMPACT_MIN,IMPACT_AVE_MIN,IMPACT_RATE_MIN,IMPACT_RATE_MAX,IMPACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� - ������� Input Check
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
        sMsg = sMsg + "������� - ȡ�������������" & Chr(13) & Chr(10)
    End If
    
    If txt_IMPACT_KND.Text = "" Then
        txt_IMPACT_KND.BackColor = &HC0E0FF
        sMsg = sMsg + "������� - ȱ�����ʹ����������" & Chr(13) & Chr(10)
    End If
    
    If txt_IMPACT_DIR.Text = "" Then
        txt_IMPACT_DIR.BackColor = &HC0E0FF
        sMsg = sMsg + "������� - �����������������" & Chr(13) & Chr(10)
    End If
    
    If iChk = 0 Then
        sdb_IMPACT_MIN.BackColor = &HC0E0FF
        sMsg = sMsg + "������� - ���ޱ�������" & Chr(13) & Chr(10)
    End If
            
    If txt_IMPACT_DSC_CD.Text = "" Then
        txt_IMPACT_DSC_CD.BackColor = &HC0E0FF
        sMsg = sMsg + "������� - �ж������������" & Chr(13) & Chr(10)
    End If
        
   'If sMsg <> "" Then MsgBox sMsg
        
    Exit Function

ResultTrue:
        GF_MATR_IMPACT_INPUT_CHECK = True
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_TIM_IMPACT_INPUT_CHECK
'   2.Name         : ���� - ʱЧ������� Input Check
'   3.Input  Value : TIM_IMPACT_SMP_CD,TIM_IMPACT_KND,TIM_IMPACT_DIR,TIM_IMPACT_MIN,TIM_IMPACT_AVE_MIN,TIM_IMPACT_RATE_MIN,TIM_IMPACT_RATE_MAX,TIM_IMPACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� - ʱЧ������� Input Check
'---------------------------------------------------------------------------------------
Public Function GF_MATR_TIM_IMPACT_INPUT_CHECK(ByVal txt_TIM_IMPACT_SMP_CD As Object, ByVal txt_TIM_IMPACT_KND As Object, ByVal txt_TIM_IMPACT_DIR As Object, ByVal sdb_TIM_IMPACT_TIM As Object, ByVal sdb_TIM_IMPACT_MIN As Object, ByVal sdb_TIM_IMPACT_MIN_MIN As Object, ByVal sdb_TIM_IMPACT_AVE_MIN As Object, ByVal sdb_TIM_IMPACT_RATE_MIN As Object, ByVal sdb_TIM_IMPACT_RATE_MAX As Object, ByVal txt_TIM_IMPACT_DSC_CD As Object) As Boolean

    Dim iCnt As Integer
    Dim iChk As Integer
    Dim sMsg As String
    Dim bCheck As Boolean
    
    txt_TIM_IMPACT_SMP_CD.BackColor = vbWhite
    txt_TIM_IMPACT_KND.BackColor = vbWhite
    txt_TIM_IMPACT_DIR.BackColor = vbWhite
    sdb_TIM_IMPACT_TIM.BackColor = vbWhite
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
    
    If sdb_TIM_IMPACT_TIM.Value <> 0 Then iCnt = iCnt + 1
    
    
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
            
    If iCnt = 0 Or (iCnt = 6 And bCheck = True) Then GoTo ResultTrue
    
    If txt_TIM_IMPACT_SMP_CD.Text = "" Then
        txt_TIM_IMPACT_SMP_CD.BackColor = &HC0E0FF
        sMsg = sMsg + "ʱЧ������� - ȡ�������������" & Chr(13) & Chr(10)
    End If
    
    If txt_TIM_IMPACT_KND.Text = "" Then
        txt_TIM_IMPACT_KND.BackColor = &HC0E0FF
        sMsg = sMsg + "ʱЧ������� - ȱ�����ʹ����������" & Chr(13) & Chr(10)
    End If
    
    If txt_TIM_IMPACT_DIR.Text = "" Then
        txt_TIM_IMPACT_DIR.BackColor = &HC0E0FF
        sMsg = sMsg + "ʱЧ������� - �����������������" & Chr(13) & Chr(10)
    End If
    
    If sdb_TIM_IMPACT_TIM.Value = 0 Then
        sdb_TIM_IMPACT_TIM.BackColor = &HC0E0FF
        sMsg = sMsg + "ʱЧ������� - ʱЧʱ���������" & Chr(13) & Chr(10)
    End If
        
    If iChk = 0 Then
        sdb_TIM_IMPACT_MIN.BackColor = &HC0E0FF
        sMsg = sMsg + "ʱЧ������� - ����ֵ��������" & Chr(13) & Chr(10)
    End If
            
    If txt_TIM_IMPACT_DSC_CD.Text = "" Then
        txt_TIM_IMPACT_DSC_CD.BackColor = &HC0E0FF
        sMsg = sMsg + "ʱЧ������� - �ж������������" & Chr(13) & Chr(10)
    End If
        
   'If sMsg <> "" Then MsgBox sMsg
        
    Exit Function

ResultTrue:
        GF_MATR_TIM_IMPACT_INPUT_CHECK = True
End Function


'---------------------------------------------------------------------------------------
'   1.ID           : GF_MATR_HIC_INPUT_CHECK
'   2.Name         : ���� - ���������� Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� - ���������� Input Check
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
'   2.Name         : ���� - ������� Input Check
'   3.Input  Value : Input Object 1 , Input Object 2 , Input Object 3 , Input Object 4 , Input Object 5 , Input Object 6
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� - ������� Input Check
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
'   2.Name         : ���� - �Ͽڼ��� Input Check
'   3.Input  Value : FRACT_SMP_CD , FRACT_KND , txt_FRACT_NAME_CD1 ,txt_FRACT_GRD1 , txt_FRACT_NAME_CD2 , txt_FRACT_GRD2 , txt_FRACT_NAME_CD3 ,txt_FRACT_GRD3 , txt_FRACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� - �Ͽڼ��� Input Check
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
'   2.Name         : ���� - �ǽ������� Input Check
'   3.Input  Value : FRACT_SMP_CD , FRACT_KND , txt_FRACT_NAME_CD1 ,txt_FRACT_GRD1 , txt_FRACT_NAME_CD2 , txt_FRACT_GRD2 , txt_FRACT_NAME_CD3 ,txt_FRACT_GRD3 , txt_FRACT_DSC_CD
'   4.Return Value : Boolean
'   5.Writer       : CHU KYO SU
'   6.Create Date  : 2003. 09 .25
'   7.Modify Date  :
'   8.Comment      : ���� - �ǽ������� Input Check
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
    Dim J As Integer
    
    J = InStr(1, sVal, ".", 1)
    
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
        
        'Call GP_SetRowHeaderClear(ss1, iRow)
        
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
Public Sub GP_ROW_PASTE(Sc As Collection, ByVal iRow As Long, Optional Mc As Collection)
    
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
    
    If Not Mc Is Nothing Then
    
        Call Gp_Ms_ControlLock(Mc("pcontrol"), False)
        
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
Public Sub GP_ROW_BACKCOLOR(ByVal ss1 As vaSpread)
    
    With ss1
        
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
'* Description   : �Է� Data�� Null���ڸ� ""�� ġȯ �� ��������,
'*                 Ư������ ����, ������ Ÿ�� ��ȯ
'* Parameter :
'*          Input :Rsdata As Variant                             -> �Է��ڷ�
'*                 DefaultValueCondtion As Variant               -> �ڷᰡ Null�ϰ�� ġȯ ���ڰ�
'*                 DataFormatTypeCondition As String             -> �������İ�
'*                 DatainSymbolTextAllDeleteCondition As String  -> ���� �� ���ڰ�
'*                 DataTypeSettingCondition As Integer           -> ��ȯ �� Ÿ���� ���� ���ڰ�
'*                 (DataType -> ������  1-String, 2-Integer, 3-Long, 4-Double)
'*
'*          ��뿹��      nullcheck("�⺻��","null�� ġȯ��","��������","Ư����ȣ����",Ÿ�Ժ�ȯ)
'*     ----------------------------------------------------------------------------
'*     2002/03/01     ������[00]    Initial Coding
'*************************************************************************************
Public Function NullCheck(Rsdata As Variant, _
                          Optional DefaultValueCondtion As Variant = "", _
                          Optional DataFormatTypeCondition As String = "", _
                          Optional DatainSymbolTextAllDeleteCondition As String = "", _
                          Optional DataTypeSettingCondition As Integer = 0) As Variant

Dim Dv As Variant                   '�⺻ ���� ����
Dim Df As String                    '����Ÿ ���� ���� ����
Dim Ds As String                    '����Ÿ���� ���� �� Ư����ȣ
Dim Dt As Variant                   '����Ÿ ��ȯ�� Ÿ�Ա��� ����

    Dv = DefaultValueCondtion
    Df = DataFormatTypeCondition
    Ds = DatainSymbolTextAllDeleteCondition
    Dt = DataTypeSettingCondition
    
    'Null���� ����
    NullCheck = IIf(IsNull(Rsdata) Or Rsdata = "", Dv, Trim(Rsdata))
    
    '����Ÿ�� ����� ���� ��ȣ ����
    If Ds <> "" Then
        NullCheck = SymbolTxtpart2(NullCheck, Ds)
    End If
    
    '����Ÿ ����� ���� ������ Formating
    If Df <> "" Then
        NullCheck = Trim(NullCheck)
        NullCheck = Format(NullCheck, Df)
    End If
    
    '���ڿ��� DataType�� ��ȯ��
    'DataType -> ������  1-String, 2-Integer, 3-Long, 4-Double
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
    ' Function �� : SymbolTxtPart2                                              '
    ' ��       �� : ���ڿ������� �߰� ��ȣ�� �����Ͽ� �����ϸ�,                 '
    '               ������ �ڷḦ �Լ��� �� ���� �մϴ�.                      '
    ' In Message  : MainTextData As String- ���ڿ�����                          '
    '               ConditionSymbol As String - ���б�ȣ                        '
    ' OutMessage  : FnMessage - ���� �� True,false�� ��ȯ Ÿ�� Variant          '                                          '                                                '
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
'   2.Name         : Check Min��Max��Tgt of Master
'   3.Input  Value : Object(dMin��dMax��dTgt)
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 07 .10
'   7.Modify Date  :
'   8.Comment      : Check Min��Max��Tgt
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
        Call Gp_MsgBoxDisplay("����У��������������Сֵ�����ֵ��Ŀ��ֵ����Ŀ", "I")
        dMin.SetFocus
    End If
                
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : GF_Sp_subValueCheck
'   2.Name         : Check Min��Max��Tgt of Master
'   3.Input  Value : Collection(SC),Long(iRow,iMin,iMax,iTgt),Object(dSetobject),String(sCMsg)
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2004. 07 .10
'   7.Modify Date  :
'   8.Comment      : Check Min��Max��Tgt
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
            Call Gp_MsgBoxDisplay("����У�����-����" + sCMsg + "����Сֵ�����ֵ��Ŀ��ֵ�����Ƿ���ȷ!", "I")
        Else
            Call Gp_MsgBoxDisplay("����У�����-�����괦����Сֵ�����ֵ��Ŀ��ֵ�����Ƿ���ȷ!", "I")
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
            Call Gp_MsgBoxDisplay("����У�����-����" + sCMsg + "�����Ƿ�����!", "I")
        Else
            Call Gp_MsgBoxDisplay("����У�����-���ڹ�괦��������!", "I")
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
Public Function GF_Sp_Necessary_Value_Check(ByVal Sc As Collection, ByVal iRow As Long, ByVal iCol As Long, Optional sCMsg As String, Optional dSetObject As Object) As Boolean
        
    GF_Sp_Necessary_Value_Check = True
    
    Sc.Item("Spread").Row = iRow
    Sc.Item("Spread").Col = iCol
    
    If Sc.Item("Spread").Text = "" Or Len(Trim(Sc.Item("Spread").Text)) = 0 Or Sc.Item("Spread").Value = 0 Then
        GF_Sp_Necessary_Value_Check = False
    End If

    If GF_Sp_Necessary_Value_Check = False Then
        
        If Not (Trim(sCMsg) = "") Then
            Call Gp_MsgBoxDisplay("����У�����-����" + sCMsg + "�����Ƿ�����!", "I")
        Else
            Call Gp_MsgBoxDisplay("����У�����-���ڹ�괦��������!", "I")
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


