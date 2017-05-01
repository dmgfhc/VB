Attribute VB_Name = "MillCommon"
Option Explicit

Public bf_OPT_LF1 As Boolean
Public af_OPT_LF1 As Boolean

Public bf_CHK_VD As Boolean
Public af_CHK_VD As Boolean
Public bf_CHK_RH As Boolean
Public af_CHK_RH As Boolean

Public sShiftSet As String
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Clear_Collection
'   2.Name         : SSCheck Control Clear Collection
'   3.Input  Value : Mc Collection
'   4.Return Value :
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 07 .27
'   7.Modify Date  :
'   8.Comment      : SSCheck Control Clear Collection
'---------------------------------------------------------------------------------------
Public Sub Gp_Clear_Collection(Name As Variant, clr As String, sControl As Collection)
    
    If LCase(Trim(clr)) = "s" Then     'Clear Key Control
        sControl.Add Item:=Name
    End If

End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_SSCheck_Cls
'   2.Name         : SSCheck Control Clear Setting
'   3.Input  Value : Mc Collection
'   4.Return Value :
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 07 .27
'   7.Modify Date  :
'   8.Comment      : SSCheck Control Clear Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_SSCheck_Cls(MC As Collection)

    Dim Ct As Control
    
    For Each Ct In MC
        If TypeOf Ct Is CheckBox Then               'CHECK BOX
            Ct.Value = UNCHECKED
            Ct.ForeColor = &H80000012
        ElseIf TypeOf Ct Is SSCheck Then            '3D CHECK BOX
            Ct.Value = ssCBUnchecked
            Ct.ForeColor = &H80000012
        End If
    Next Ct

End Sub

'
'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Process
'   2.Name         : Spread Data Process
'   3.Input  Value : Conn Connection, Sc Collection, Mc Collection, {RefChek,MAT_CD Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 07 .27
'   7.Modify Date  :
'   8.Comment      : Spread Data Process
'---------------------------------------------------------------------------------------
Public Function Gf_Mill_Process(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False, Optional MAT_CD As String) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim sMesg As String
    Dim sTemp As String
    Dim ProcessChk As String
    Dim DelYN As Boolean
    Dim Msg_Count As Integer
    Dim Msg_Yes As String
    
    Dim adoCmd As ADODB.Command

    Gf_Mill_Process = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
        Gf_Mill_Process = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(iCount, Sc, MC)
                End If
                
                'Maxlength Check
                sMesg = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "长度不正确"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Mill_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Mill_Process = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Mill_Process = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To Sc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    Msg_Count = 1
    For iCount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        DelYN = False
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input"
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).Value = "D"
                ProcessChk = "YES"
        End Select
          
        If ProcessChk = "YES" Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempFloat)
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempInt)
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Text = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Messg Check
            
            If adoCmd("Error") = "2" Then
            
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = ret_Result_ErrMsg
                Call Gp_MsgBoxDisplay(sErrMessg, "I")
                
            End If
            
            'Error Check
            
            If adoCmd("Error") = "1" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Mill_Process = False
                Exit Function
             
             End If
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input", "Update"
                Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                
            Case "Delete"
                If DelYN Then
                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), iCount)
                   iCount = iCount - 1
                End If
        End Select
        
    Next iCount
    
    Sc.Item("Spread").ReDraw = True
    
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            If RefChek = False Then Gf_Mill_Process = Gf_Sp_Display(Conn, Sc.Item("Spread"), _
                                                    Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", MC("pControl")), Sc.Item("pColumn"), False)
        Else
            If RefChek = False Then Gf_Mill_Process = Gf_Sp_Display(Conn, Sc.Item("Spread"), _
                           Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), False)
        End If
        
        If MAT_CD = "" Then
           MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了 " & iProcessCount & " 条记录"
        ElseIf MAT_CD = "P" Then
           MDIMain.StatusBar1.Panels(1) = "提示信息： " & iProcessCount & " 块钢板完成了入库/倒库操作"
            'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
        ElseIf MAT_CD = "C" Then
           MDIMain.StatusBar1.Panels(1) = "提示信息： " & iProcessCount & " 个钢卷完成了入库/倒库操作"
        End If
        
    End If
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Gf_Mill_Process = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Mill_Process = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function
'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Plate_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant, sQuery String,, {RefChk,ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : Combo Add
'---------------------------------------------------------------------------------------
Public Function Gf_Plate_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sQuery As String, _
                  Optional RefChk As Boolean = True, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim AdoRs As ADODB.Recordset
    Dim sPlate_no As String
    Dim sLast_Plate_no As String
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Plate_ComboAdd = False: Exit Function
    End If
        
    If RefChk = True Then
       sPlate_no = ""
    Else
       sPlate_no = Cbo.Text
    End If
    
    If ClsChk Then
       Cbo.Clear
    End If

    Cbo.Text = sPlate_no
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0).Value) <> vbNull And AdoRs.Fields(1).Value = "C" Then
                Cbo.AddItem AdoRs.Fields(0).Value + "->"
            ElseIf VarType(AdoRs.Fields(0).Value) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0).Value
                If sPlate_no = "" Then
                   sPlate_no = AdoRs.Fields(0).Value
                   Cbo.Text = sPlate_no
                End If
            End If
            
            sLast_Plate_no = AdoRs.Fields(0).Value
            
            AdoRs.MoveNext
            
        Wend
        Gf_Plate_ComboAdd = True
    Else
        Gf_Plate_ComboAdd = False
    End If

    If sPlate_no = "" Then
       Cbo.Text = sLast_Plate_no
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:
'
    Set AdoRs = Nothing
    Gf_Plate_ComboAdd = False

End Function
'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Plate_ComboSet
'   2.Name         :
'   3.Input  Value : Conn Connection, MC Collection, Cbo Variant, {RefChk,ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : Combo Set
'---------------------------------------------------------------------------------------
Public Function Gf_Plate_ComboSet(Conn As ADODB.Connection, MC As Collection, Cbo As Variant, _
                 Optional RefChk As Boolean = True, Optional ClsChk As Boolean = True) As Boolean
    
On Error GoTo MasterRef_Err

    Dim sQuery As String

    'Make Query
    sQuery = Gf_Ms_MakeQuery(MC.Item("P-R"), "R", MC.Item("pControl"))
    
    If sQuery = "FAIL" Then
        Exit Function
    End If
    
    'Query Excete and Display
    If Gf_Plate_ComboAdd(M_CN1, Cbo, sQuery, RefChk, ClsChk) Then
        Gf_Plate_ComboSet = True
        Exit Function
    End If
    
    Gf_Plate_ComboSet = False
    Exit Function

MasterRef_Err:

'    Call Gp_MsgBoxDisplay("Failed on data inquiry")
    Gf_Plate_ComboSet = False

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_DTSet
'   2.Name         : Get System/Vb Date,Time
'   3.Input  Value : Conn Connection, {DTCheck,DTFlag String}
'   4.Return Value : Variant
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .24
'   7.Modify Date  :
'   8.Comment      : Get System/Vb Date,Time
'---------------------------------------------------------------------------------------
Public Function Gf_DTSet(Conn As ADODB.Connection, Optional DTCheck As String = "S", Optional DTFlag As String = "C") As Variant

On Error GoTo DTSet_Error

    Dim sQuery As String
    Dim sQuery_Len As Long
    
    Select Case DTCheck
           Case "S"
           sQuery = "SELECT TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS') FROM DUAL"
           sQuery_Len = 14
           Case "I"
           sQuery = "SELECT TO_CHAR(SYSDATE,'YYYYMMDDHH24MI') FROM DUAL"
           sQuery_Len = 12
           Case "H"
           sQuery = "SELECT TO_CHAR(SYSDATE,'YYYYMMDDHH24') FROM DUAL"
           sQuery_Len = 10
           Case "D"
           sQuery = "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL"
           sQuery_Len = 8
           Case "T"
           sQuery = "SELECT TO_CHAR(SYSDATE,'HH24MISS') FROM DUAL"
           sQuery_Len = 6
           Case "M"
           sQuery = "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL"
           sQuery_Len = 6
           Case "Y"
           sQuery = "SELECT TO_CHAR(SYSDATE,'YYYY') FROM DUAL"
           sQuery_Len = 4
    End Select
    
    If DTFlag = "C" Then
       Gf_DTSet = Mid(Format(Now, "YYYYMMDDHHMMSS"), 1, sQuery_Len)
       If DTCheck = "T" Then
          Gf_DTSet = Format(Now, "HHMMSS")
       End If
       Exit Function
    End If
       
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_DTSet = "00000000000000": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_DTSet = ""
            Else
                Gf_DTSet = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_DTSet = "00000000000000"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

DTSet_Error:

    Set AdoRs = Nothing
    Gf_DTSet = "00000000000000"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_ShiftSet
'   2.Name         : Shift Return
'   3.Input  Value : Conn Connection
'   4.Return Value : Variant
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .24
'   7.Modify Date  :
'   8.Comment      : Shift Return
'---------------------------------------------------------------------------------------
Public Function Gf_ShiftSet(Conn As ADODB.Connection, Optional WKDATE As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim Shift_HH As String
    Dim AdoRs As ADODB.Recordset
    
    If WKDATE = "" Then
        sQuery = "SELECT TO_CHAR(SYSDATE,'HH24MI') FROM DUAL"
    
        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Gf_ShiftSet = "0": Exit Function
        End If
        
        Set AdoRs = New ADODB.Recordset
    
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If Not AdoRs.BOF And Not AdoRs.EOF Then
        
            If Not AdoRs.EOF Then
                If VarType(AdoRs.Fields(0)) = vbNull Then
                    Shift_HH = ""
                Else
                    Shift_HH = AdoRs.Fields(0)
                End If
            End If
            
        Else
            Shift_HH = ""
        End If
        
        AdoRs.Close
        Set AdoRs = Nothing
    Else
        Shift_HH = WKDATE

    End If
    
    If Val(Shift_HH) < 800 Then
       Gf_ShiftSet = "1"
    ElseIf Val(Shift_HH) < 1600 Then
        Gf_ShiftSet = "2"
    Else
        Gf_ShiftSet = "3"
    End If
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_ShiftSet = "0"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_ShiftSet
'   2.Name         : Shift Return
'   3.Input  Value : Conn Connection
'   4.Return Value : Variant
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .24
'   7.Modify Date  :
'   8.Comment      : Shift Return
'---------------------------------------------------------------------------------------
Public Function Gf_ShiftSet3(Conn As ADODB.Connection, Optional WKDATE As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim Shift_HH As String
    Dim AdoRs As ADODB.Recordset
    
    If WKDATE = "" Then
        sQuery = "SELECT TO_CHAR(SYSDATE,'HH24MI') FROM DUAL"
    
        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Gf_ShiftSet3 = "0": Exit Function
        End If
        
        Set AdoRs = New ADODB.Recordset
    
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If Not AdoRs.BOF And Not AdoRs.EOF Then
        
            If Not AdoRs.EOF Then
                If VarType(AdoRs.Fields(0)) = vbNull Then
                    Shift_HH = ""
                Else
                    Shift_HH = AdoRs.Fields(0)
                End If
            End If
            
        Else
            Shift_HH = ""
        End If
    Else
        Shift_HH = WKDATE
    End If
    
    If Val(Shift_HH) < 800 Then
       Gf_ShiftSet3 = "1"
    ElseIf Val(Shift_HH) < 1600 Then
        Gf_ShiftSet3 = "2"
    Else
        Gf_ShiftSet3 = "3"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_ShiftSet3 = "0"

End Function
Public Function Gf_GroupSet(Conn As ADODB.Connection, Shift As String, setDate) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim Shift_HH As String
    Dim AdoRs As ADODB.Recordset
    
    sQuery = "SELECT Gf_Groupset('C3'," & Shift & ",SUBSTR('" & setDate & "',1,8)) FROM DUAL"

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_GroupSet = "0": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_GroupSet = ""
            Else
                Gf_GroupSet = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_GroupSet = ""
    End If
     
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_GroupSet = "0"

End Function


'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Plate_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant,sPRC String,
'                    {sFACT_CD,sPRC_LINE String, sADDNUM As Integer, ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : combo Add
'---------------------------------------------------------------------------------------
Public Function Gf_Mill_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPrc As String, Optional sFACT_CD As String = "C1", _
             Optional sPRC_LINE As String = "1", Optional sADDNUM As Integer = 20, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String

    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Mill_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT GOODS_ID FROM (SELECT B.GOODS_ID "
    sQuery = sQuery + "               FROM FP_TRACKIDX A, FP_TRACKDATA B "
    sQuery = sQuery + "              WHERE A.FACT_CD = '" + sFACT_CD + "'"
    sQuery = sQuery + "                AND A.PRC = '" + sPrc + "'"
    sQuery = sQuery + "                AND A.PRC_LINE= '" + sPRC_LINE + "'"
    sQuery = sQuery + "                AND A.FACT_CD=B.FACT_CD "
    sQuery = sQuery + "                AND A.PRC=B.PRC "
    sQuery = sQuery + "                AND A.PRC_LINE=B.PRC_LINE "
    sQuery = sQuery + "                AND B.SEQ_NO <= A.LAST_SEQ "
    sQuery = sQuery + "           ORDER BY B.SEQ_NO DESC) "
    sQuery = sQuery + "              WHERE ROWNUM <= " + CStr(sADDNUM)

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_Mill_ComboAdd = True
    Else
        Gf_Mill_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_Mill_ComboAdd = False

End Function
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_MsgBox
'   2.Name         : Message Box Display
'   3.Input  Value : sMsg String, {sIcon String}, {sTitle String}
'   4.Return Value :
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : Message Box
'---------------------------------------------------------------------------------------
Public Function Gp_MsgBox(ByVal sMsg As String, Optional sIcon As String, Optional sTitle As String) As Integer
    
    Dim sStyle As String
    Dim iRet As Integer
    
    'Message Box Style Selection
    Select Case sIcon
        Case "Q"
            sStyle = vbOKOnly + vbQuestion
        Case "W"
            sStyle = vbOKOnly + vbExclamation
        Case "I"
            sStyle = vbOKOnly + vbInformation
        Case "C"
            sStyle = vbYesNo + vbQuestion
        Case Else
            sStyle = vbOKOnly + vbCritical
    End Select
    
    If RTrim(sTitle) = "" Then
        sTitle = "系统提示信息"
    End If

    If MsgBox(sMsg, sStyle, sTitle) = vbYes Then
       Gp_MsgBox = 6
    Else
       Gp_MsgBox = 7
    End If
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_DateCheck
'   2.Name         : DateTime Check
'   3.Input  Value : DateCheck As Variant
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .27
'   7.Modify Date  :
'   8.Comment      : Date,Time Check
'---------------------------------------------------------------------------------------
Public Function Gp_DateCheck(DateCheck As Variant, Optional sDTChk As String = "M") As Boolean

On Error GoTo DateCheck_Error

    Dim iDateCheck As String
    Dim iDateMatch As String
    Dim iDate As String
    Dim iCheck As Date
    
    If sDTChk = "M" Then
       iDateCheck = DateCheck.RawData
    Else
       iDateCheck = Replace(DateCheck, "-", "")
       iDateCheck = Replace(iDateCheck, " ", "")
       iDateCheck = Replace(iDateCheck, ":", "")
    End If
    
    If Val(Mid(iDateCheck, 1, 4)) > 2020 Or Val(Mid(iDateCheck, 1, 4)) < 2000 Then
       Gp_DateCheck = False
       Exit Function
    End If
       
    Select Case Len(iDateCheck)
           Case 8
                iDate = Mid(iDateCheck, 1, 4) + "-" + Mid(iDateCheck, 5, 2) + "-" + Mid(iDateCheck, 7, 2)
                iCheck = CDate(Mid(iDate, 2, 10))
           Case 12
                iDate = Mid(iDateCheck, 1, 4) + "-" + Mid(iDateCheck, 5, 2) + "-" + Mid(iDateCheck, 7, 2) _
                + " " + Mid(iDateCheck, 9, 2) + ":" + Mid(iDateCheck, 11, 2)
                iCheck = CDate(Mid(iDate, 2, 16))
           Case 14
                iDate = Mid(iDateCheck, 1, 4) + "-" + Mid(iDateCheck, 5, 2) + "-" + Mid(iDateCheck, 7, 2) _
                + " " + Mid(iDateCheck, 9, 2) + ":" + Mid(iDateCheck, 11, 2) + ":" + Mid(iDateCheck, 13, 2)
                iCheck = CDate(Mid(iDate, 2, 19))
           Case Else
                Gp_DateCheck = False
                Exit Function
    End Select
    
    iDateMatch = Format(iCheck, "YYYYMMDD")
    
    If iDateMatch <> Mid(iDateCheck, 1, 8) Then
        Gp_DateCheck = False
        Exit Function
    End If
    
    Gp_DateCheck = True
    Exit Function

DateCheck_Error:

    Gp_DateCheck = False
    Exit Function

End Function
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Mill_ControlLock
'   2.Name         : Control Lock
'   3.Input  Value : lControl Collection, Tf Boolean, {EL As String}
'   4.Return Value :
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 09. 09
'   7.Modify Date  :
'   8.Comment      : Control Lock
'---------------------------------------------------------------------------------------
Public Sub Gp_Mill_ControlLock(lControl As Collection, Tf As Boolean, Optional EL As String = "L")
    
    Dim iCount As Integer
     
    If lControl.Count < 1 Then Exit Sub
    
    If EL = "E" Then
    
       For iCount = 1 To lControl.Count
           lControl.Item(iCount).Enabled = Not Tf
       Next iCount
       
    Else
        
       For iCount = 1 To lControl.Count
        
            If TypeOf lControl.Item(iCount) Is ComboBox Then           'COMBO BOX
                   lControl.Item(iCount).Locked = Tf
            ElseIf TypeOf lControl.Item(iCount) Is TextBox Then        'TextBox
                   lControl.Item(iCount).Locked = Tf
            ElseIf TypeOf lControl.Item(iCount) Is sidbEdit Then       'sidbEdit
                   lControl.Item(iCount).ReadOnly = Tf
            ElseIf TypeOf lControl.Item(iCount) Is silgEdit Then       'silgEdit
                   lControl.Item(iCount).ReadOnly = Tf
            ElseIf TypeOf lControl.Item(iCount) Is sidtEdit Then       'sidtEdit
                   lControl.Item(iCount).ReadOnly = Tf
            ElseIf TypeOf lControl.Item(iCount) Is sitmEdit Then       'sitmEdit
                   lControl.Item(iCount).ReadOnly = Tf
            ElseIf TypeOf lControl.Item(iCount) Is sitxEdit Then       'sitxEdit
                   lControl.Item(iCount).ReadOnly = Tf
            Else
                   lControl.Item(iCount).Enabled = Not Tf
            End If
            
       Next iCount
       
     End If
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Common_ComboSet
'   2.Name         :
'   3.Input  Value : Conn Connection, MC Collection, Cbo Variant, {RefChk,ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : Combo Set
'---------------------------------------------------------------------------------------
Public Function Gf_Common_ComboSet(Conn As ADODB.Connection, Cbo As Variant, _
                                   sPrc As String, Optional ClsChk As Boolean = True) As Boolean
    
On Error GoTo MasterRef_Err

    Dim sQuery As String

    'Make Query
    sQuery = "{call AGT_CBOSET.P_REFER ('" + "C1" + "','" + sPrc + "','" + "1" + "')}" ' "',?)}"
    
    If sQuery = "FAIL" Then
        Exit Function
    End If
    
    'Query Excete and Display
    If Gf_Common_ComboAdd(M_CN1, Cbo, sQuery, ClsChk) Then
        Gf_Common_ComboSet = True
        Exit Function
    End If
    
    Gf_Common_ComboSet = False
    Exit Function

MasterRef_Err:

'    Call Gp_MsgBoxDisplay("Failed on data inquiry")
    Gf_Common_ComboSet = False

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Common_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant, sQuery String,, {RefChk,ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : Combo Add
'---------------------------------------------------------------------------------------
Public Function Gf_Common_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sQuery As String, _
                                   Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim AdoRs As ADODB.Recordset
'    Dim sPlate_no As String
'    Dim sLast_Plate_no As String
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Common_ComboAdd = False: Exit Function
    End If
  
    If ClsChk Then
       Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0).Value) <> vbNull And AdoRs.Fields(1).Value <> "A" Then
                Cbo.AddItem AdoRs.Fields(0).Value + "->"
            ElseIf VarType(AdoRs.Fields(0).Value) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0).Value
            End If
            
            AdoRs.MoveNext
            
        Wend
        Gf_Common_ComboAdd = True
    Else
        Gf_Common_ComboAdd = False
    End If

    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:
'
    Set AdoRs = Nothing
    Gf_Common_ComboAdd = False

End Function
'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Mill_Common_DD
'   2.Name         : Common Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2005. 10 .15
'   7.Modify Date  :
'   8.Comment      : Common Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Mill_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
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
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Or DD.nameType = "" Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
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
    
    DD.DataDicType = "M"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT CD ""代码"", CD_SHORT_NAME ""代码简称"", CD_NAME ""代码名称"", "
        DD.sQuery = DD.sQuery + "       CD_SHORT_ENG ""代码英文简称"", CD_FULL_ENG ""代码英文名称"" FROM NISCO.ZP_CD "
        DD.sQuery = DD.sQuery + " WHERE CD_MANA_NO        =    '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND CD                like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND NVL(APLY_STD,'N') =    'Y' "
        
        If DD.rControl.Count > 1 Then
            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_SHORT_NAME,'%') like '" & Trim(DD.rControl.Item(2).Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_NAME,'%')       like '" & Trim(DD.rControl.Item(2).Text) & "%' "
                Case "3"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_SHORT_ENG,'%')  like '" & Trim(DD.rControl.Item(2).Text) & "%' "
                Case "4"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_FULL_ENG,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
            End Select
        End If
    
    Else

        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text

        DD.sQuery = "            SELECT CD ""代码"", CD_SHORT_NAME ""代码简称"", CD_NAME ""代码名称"", "
        DD.sQuery = DD.sQuery + "       CD_SHORT_ENG ""代码英文简称"", CD_FULL_ENG ""代码英文名称"" FROM NISCO.ZP_CD "
        DD.sQuery = DD.sQuery + " WHERE CD_MANA_NO =    '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND CD         like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND NVL(APLY_STD,'N') =    'Y' "

        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text

            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_SHORT_NAME,'%') like '" & Trim(DD.sPname.Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_NAME,'%')       like '" & Trim(DD.sPname.Text) & "%' "
                Case "3"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_SHORT_ENG,'%')  like '" & Trim(DD.sPname.Text) & "%' "
                Case "4"
                    DD.sWhere = DD.sWhere + " AND NVL(CD_FULL_ENG,'%')   like '" & Trim(DD.sPname.Text) & "%' "
            End Select
        End If

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
'   1.ID           : Gf_HeatNo_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant, sTableName String, sColId String, sPrcLine String, {ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim S.H
'   6.Create Date  : 2006. 01 .24
'   7.Modify Date  :
'   8.Comment      : Add Heat No in combo box
'---------------------------------------------------------------------------------------
Public Function Gf_HeatNo_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, _
                                   sTableName As String, sColId As String, sPrcLine As String, _
                                   Optional ClsChk As Boolean = True) As Boolean

On Error GoTo Gf_HeatNo_ComboAdd_Error
    
    Dim AdoRs  As ADODB.Recordset
    Dim sQuery As String
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_HeatNo_ComboAdd = False: Exit Function
    End If
    
    If ClsChk Then
        Cbo.Clear
    End If
     
    sQuery = "SELECT C.HEAT_MANA_NO, C." & sColId
    sQuery = sQuery & "  FROM ( "
    sQuery = sQuery & "          SELECT A.HEAT_MANA_NO,B." & sColId
    sQuery = sQuery & "            FROM EP_CHARGE_IDX A, " & sTableName & " B"
    sQuery = sQuery & "           WHERE A.PRC_STS      IN  ('A','B') "
    sQuery = sQuery & "             AND A.PRC_LINE     LIKE '" & Trim(sPrcLine) & "%'"
    sQuery = sQuery & "             AND A.HEAT_MANA_NO =  B.HEAT_NO(+)"
    sQuery = sQuery & "           ORDER BY A.HEAT_MANA_NO) C"
    sQuery = sQuery & "   WHERE ROWNUM <= 15 "
    
    Set AdoRs = New ADODB.Recordset
     
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
     
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
               If AdoRs.Fields(1) > "0" Then
                  
                  Cbo.AddItem AdoRs.Fields(0)
                  
               Else
                  Cbo.AddItem AdoRs.Fields(0)
              
              
               End If
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_HeatNo_ComboAdd = True
    Else
        Gf_HeatNo_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

Gf_HeatNo_ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_HeatNo_ComboAdd = False

End Function
