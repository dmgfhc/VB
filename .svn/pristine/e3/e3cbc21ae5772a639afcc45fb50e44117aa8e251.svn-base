Attribute VB_Name = "ReferCommon"
Option Explicit

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Ms_RefCollection
'   2.Name         : Master Refer Collection Setting
'   3.Input  Value : Name Variant, rctl String, rControl Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Refer Collection Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Ms_RefCollection(Name As Variant, rctl As String, rControl As Collection)
    
    If LCase(Trim(rctl)) = "r" Then     'Refer Control
        rControl.Add Item:=Name
    End If
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Only_Display
'   2.Name         : Only Display
'   3.Input  Value : Conn Connection, Sc Collection, sQuery String, {iDupCnt Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Only Display
'---------------------------------------------------------------------------------------
Public Function Gf_Only_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional iDupCnt As Variant = 0, _
                                Optional MsgChk As Boolean = True, Optional EvenRowChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn
    
    Dim JJ, j As Integer
    
    Dim lRowCount As Long
    Dim lColCount As Long
    Dim sTemp() As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Only_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
        
    With Sc.Item("Spread")

        Gf_Only_Display = True
        
        .ReDraw = False
        .MaxRows = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Gf_Only_Display = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = 0
            
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If iDupCnt > 0 Then
            ReDim sTemp(0 To iDupCnt - 1)
        End If
        
        If UBound(ArrayRecords, 1) >= 0 Then
        
            .MaxRows = UBound(ArrayRecords, 2) + 1
        
            For lRowCount = 0 To .MaxRows - 1
            
                .Row = lRowCount + 1
                
                'Duplicate Process
                For j = 1 To iDupCnt Step 1
                
                    If sTemp(j - 1) <> Trim(ArrayRecords(j - 1, lRowCount)) Then
                        .Col = j
                        .Text = Trim(ArrayRecords(j - 1, lRowCount))
                        sTemp(j - 1) = Trim(ArrayRecords(j - 1, lRowCount))
                        
                        For JJ = j + 1 To iDupCnt Step 1
                            sTemp(JJ - 1) = ""
                        Next JJ
                        
                    End If
                    
                Next j
            
                For lColCount = iDupCnt To .MaxCols - 1
                
                    .Col = lColCount + 1
                    
                    If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(lColCount, lRowCount))
                    End If
                    
                Next lColCount
                
            Next lRowCount
        End If
                                            
        .ReDraw = True
        
    End With
    
    If EvenRowChk Then Call Gp_Sp_EvenRowBackcolor(Sc.Item("Spread"))
    
    Sc.Item("Spread").OperationMode = OperationModeRow
    
    Gf_Only_Display = True
    Screen.MousePointer = vbDefault
    
    Exit Function
   
Error_Rtn:

    Set AdoRs = Nothing
    Gf_Only_Display = False

    Screen.MousePointer = vbDefault
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Total_Display
'   2.Name         : Total Display
'   3.Input  Value : Conn Connection, Sc Collection, sQuery String,
'                    {iDupCnt Variant}, {iSumCnt Variant}, {iSumCol Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Total Display
'---------------------------------------------------------------------------------------
Public Function Gf_Total_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional iDupCnt As Variant = 0, _
                                 Optional iSumCnt As Variant = 0, Optional iSumCol As Variant, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn
    
    Dim k As Long
    Dim JJ, j As Integer
    Dim iBas As Integer
    Dim iCot As Integer
    
    Dim lRowCount As Long
    Dim lColCount As Long
    
    Dim sCol_a As String
    Dim sCol_b As String
    Dim sTemp() As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Total_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    With Sc.Item("Spread")
    
        Gf_Total_Display = True
        
        .ReDraw = False
        .MaxRows = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Gf_Total_Display = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
        
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If iDupCnt > 0 Then
            ReDim sTemp(0 To iDupCnt - 1)
        End If
        
        If UBound(ArrayRecords, 1) <> 0 Then
        
            .MaxRows = UBound(ArrayRecords, 2) + 1
        
            For lRowCount = 0 To .MaxRows - 1
            
                .Row = lRowCount + 1
                
                'Duplicate Process
                For j = 1 To iDupCnt Step 1
                
                    If sTemp(j - 1) <> Trim(ArrayRecords(j - 1, lRowCount)) Then
                        .Col = j
                        .Text = Trim(ArrayRecords(j - 1, lRowCount))
                        sTemp(j - 1) = Trim(ArrayRecords(j - 1, lRowCount))
                        
                        For JJ = j + 1 To iDupCnt Step 1
                            sTemp(JJ - 1) = ""
                        Next JJ
                        
                    End If
                    
                Next j
            
                For lColCount = iDupCnt To .MaxCols - 1
                    
                    .Col = lColCount + 1

                    If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(lColCount, lRowCount))
                    End If
                    
                Next lColCount
                
            Next lRowCount
            
        End If
                                            
        'Total Compute
        If iSumCnt > 0 Then
        
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
             
            'Color
            '&HFFE6E6 sub total 1
            '&HE6FFE6 sub total 2
            '&HE6E6FF total
            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
            
            For j = 1 To iSumCnt
                .Col = j
                If .ColHidden = False Then
                    .Text = "合  计"
                    j = iSumCnt
                End If
            Next j
            
            For j = 1 To iSumCnt
                .Col = iSumCol(j)
                
                If iSumCol(j) <= 26 Then
                    sCol_a = Chr(iSumCol(j) + 64)
                    .Formula = "sum(" + sCol_a + "1:" + sCol_a & .MaxRows - 1 & ")"
                Else
                    iCot = Int(((iSumCol(j) - 1) / 26))
                    iBas = 26 * iCot
                    sCol_a = Chr((iSumCol(j) - iBas) + 64)
                    sCol_b = Chr(iCot + 64)
                    .Formula = "sum(" + sCol_b + sCol_a + "1:" + sCol_b + sCol_a & .MaxRows - 1 & ")"
                End If
            Next j
            
        End If
            
        .ReDraw = True
        Gf_Total_Display = True
        
    End With
    
    Call Gp_Sp_EvenRowBackcolor(Sc.Item("Spread"), 1)
    
    Sc.Item("Spread").OperationMode = OperationModeRow
    Screen.MousePointer = vbDefault
    
    Exit Function
    
Error_Rtn:

    Set AdoRs = Nothing
    Gf_Total_Display = False
    Screen.MousePointer = vbDefault
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Stotal_Display
'   2.Name         : SubTotal Display
'   3.Input  Value : Conn Connection, Sc Collection, sQuery String,
'                    {iDupCnt Variant}, {iSumCnt Variant}, {iSumCol Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : SubTotal Display
'---------------------------------------------------------------------------------------
Public Function Gf_Sonly_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional iDupCnt As Variant = 0, _
                                  Optional iSumCnt As Variant = 0, Optional iSumCol As Variant, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn

    Dim k, j As Long

    Dim iBas As Integer
    Dim iCot As Integer
    Dim iOld_Row As Integer
    
    Dim Sw As Boolean

    Dim lRowCount As Long
    Dim lColCount As Long

    Dim sCol_a As String
    Dim sCol_b As String
    Dim sTemp() As String

    Dim dTotal() As Double

    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Sonly_Display = False: Exit Function
    End If
    
    If iSumCnt <> iSumCol.Count Then
        Call Gp_MsgBoxDisplay("iSumCnt and iSumCol are different..!!", "I")
        Gf_Sonly_Display = False
        Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    ReDim sTemp(0 To iDupCnt - 1) As String
    ReDim dTotal(1 To iSumCnt) As Double
        
    With Sc.Item("Spread")
    
        Gf_Sonly_Display = True
        
        .ReDraw = False
        .MaxRows = 0
        iOld_Row = 1
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
        
            Gf_Sonly_Display = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
        
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 1) <> 0 Then
        
            For lRowCount = 0 To UBound(ArrayRecords, 2)
            
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                'Duplicate Process
                For j = 1 To iDupCnt Step 1
                
                    If sTemp(j - 1) <> Trim(ArrayRecords(j - 1, lRowCount)) Then
    
                        'Sub Total
                        '--------------------------------------------------------------------------
                        If .Row <> 1 And j = 1 Then
                            .Col = 1
    
                            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .Row, .Row, BLACK, &HFFE6E6)
                            
                            For k = 1 To iDupCnt
                                .Col = k
                                If .ColHidden = False Then
                                    .Text = "小  计"
                                    k = iDupCnt
                                End If
                            Next k
                            
                            For k = 1 To iSumCnt
                                .Col = iSumCol(k)
    
                                If iSumCol(k) <= 26 Then
                                    sCol_a = Chr(iSumCol(k) + 64)
                                    .Formula = "sum(" + sCol_a & iOld_Row & ":" + sCol_a & .Row - 1 & ")"
                                Else
                                    iCot = Int(((iSumCol(k) - 1) / 26))
                                    iBas = 26 * iCot
    
                                    sCol_a = Chr((iSumCol(k) - iBas) + 64)
                                    sCol_b = Chr(iCot + 64)
    
                                    .Formula = "sum(" + sCol_b + sCol_a & iOld_Row & ":" + sCol_b + sCol_a & .Row - 1 & ")"
                                End If
                                
                                If VarType(.Value) = vbNull Then
                                    dTotal(k) = dTotal(k) + 0
                                Else
                                    dTotal(k) = dTotal(k) + .Value
                                End If
                                
                            Next k
                            
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            iOld_Row = .Row
                            
                        End If
                        '--------------------------------------------------------------------------
                        .Col = j
                        .Text = ArrayRecords(j - 1, lRowCount)
                        sTemp(j - 1) = ArrayRecords(j - 1, lRowCount)
                        
                        For k = j + 1 To iDupCnt Step 1
                            sTemp(k - 1) = ""
                        Next k
                        
                    End If
                
                Next j

                'Duplicate Process
                For lColCount = iDupCnt To .MaxCols - 1
                    .Col = lColCount + 1
                    
                    If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(lColCount, lRowCount))
                    End If
                        
                Next lColCount
    
            Next lRowCount
            
        End If
        
        'Last Sub Total
        '--------------------------------------------------------------------------
        If .MaxRows > 0 Then
        
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1

            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HFFE6E6)
            
            For k = 1 To iDupCnt
                .Col = k
                If .ColHidden = False Then
                    .Text = "小  计"
                    k = iDupCnt
                End If
            Next k
            
            For k = 1 To iSumCnt
                .Col = iSumCol(k)

                If iSumCol(k) <= 26 Then
                    sCol_a = Chr(iSumCol(k) + 64)
                    .Formula = "sum(" + sCol_a & iOld_Row & ":" + sCol_a & .Row - 1 & ")"
                Else
                    iCot = Int(((iSumCol(k) - 1) / 26))
                    iBas = 26 * iCot

                    sCol_a = Chr((iSumCol(k) - iBas) + 64)
                    sCol_b = Chr(iCot + 64)

                    .Formula = "sum(" + sCol_b + sCol_a & iOld_Row & ":" + sCol_b + sCol_a & .Row - 1 & ")"
                End If
                
                If VarType(.Value) = vbNull Then
                    dTotal(k) = dTotal(k) + 0
                Else
                    dTotal(k) = dTotal(k) + .Value
                End If
                
            Next k
            
        End If

        .ReDraw = True
        
        Gf_Sonly_Display = True
        Sc.Item("Spread").OperationMode = OperationModeRow
        
        Screen.MousePointer = vbDefault
    End With
    
    Exit Function

Error_Rtn:

    Set AdoRs = Nothing
    Gf_Sonly_Display = False

    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Stotal_Display
'   2.Name         : Sub, Total Display
'   3.Input  Value : Conn Connection, Sc Collection, sQuery String,
'                    {iDupCnt Variant}, {iSumCnt Variant}, {iSumCol Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Sub, Total Display
'---------------------------------------------------------------------------------------
Public Function Gf_Stotal_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional iDupCnt As Variant = 0, _
                                  Optional iSumCnt As Variant = 0, Optional iSumCol As Variant, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn

    Dim k, j As Long

    Dim iBas As Integer
    Dim iCot As Integer
    Dim iOld_Row As Integer
    
    Dim Sw As Boolean

    Dim lRowCount As Long
    Dim lColCount As Long

    Dim sCol_a As String
    Dim sCol_b As String
    Dim sTemp() As String

    Dim dTotal() As Double

    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Stotal_Display = False: Exit Function
    End If
    
    If iSumCnt <> iSumCol.Count Then
        Call Gp_MsgBoxDisplay("iSumCnt and iSumCol are different..!!", "I")
        Gf_Stotal_Display = False
        Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    ReDim sTemp(0 To iDupCnt - 1) As String
    ReDim dTotal(1 To iSumCnt) As Double
        
    With Sc.Item("Spread")
    
        Gf_Stotal_Display = True
        
        .ReDraw = False
        .MaxRows = 0
        iOld_Row = 1
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
        
            Gf_Stotal_Display = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
        
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 1) <> 0 Then
        
            For lRowCount = 0 To UBound(ArrayRecords, 2)
            
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                'Duplicate Process
                For j = 1 To iDupCnt Step 1
                
                    If sTemp(j - 1) <> Trim(ArrayRecords(j - 1, lRowCount)) Then
    
                        'Sub Total
                        '--------------------------------------------------------------------------
                        If .Row <> 1 And j = 1 Then
                            .Col = 1
    
                            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .Row, .Row, BLACK, &HFFE6E6)
                            
                            For k = 1 To iDupCnt
                                .Col = k
                                If .ColHidden = False Then
                                    .Text = "小  计"
                                    k = iDupCnt
                                End If
                            Next k
                            
                            For k = 1 To iSumCnt
                                .Col = iSumCol(k)
    
                                If iSumCol(k) <= 26 Then
                                    sCol_a = Chr(iSumCol(k) + 64)
                                    .Formula = "sum(" + sCol_a & iOld_Row & ":" + sCol_a & .Row - 1 & ")"
                                Else
                                    iCot = Int(((iSumCol(k) - 1) / 26))
                                    iBas = 26 * iCot
    
                                    sCol_a = Chr((iSumCol(k) - iBas) + 64)
                                    sCol_b = Chr(iCot + 64)
    
                                    .Formula = "sum(" + sCol_b + sCol_a & iOld_Row & ":" + sCol_b + sCol_a & .Row - 1 & ")"
                                End If
                                
                                If VarType(.Value) = vbNull Then
                                    dTotal(k) = dTotal(k) + 0
                                Else
                                    dTotal(k) = dTotal(k) + .Value
                                End If
                                
                            Next k
                            
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            iOld_Row = .Row
                            
                        End If
                        '--------------------------------------------------------------------------
                        .Col = j
                        .Text = ArrayRecords(j - 1, lRowCount)
                        sTemp(j - 1) = ArrayRecords(j - 1, lRowCount)
                        
                        For k = j + 1 To iDupCnt Step 1
                            sTemp(k - 1) = ""
                        Next k
                        
                    End If
                
                Next j

                'Duplicate Process
                For lColCount = iDupCnt To .MaxCols - 1
                    .Col = lColCount + 1
                    
                    If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(lColCount, lRowCount))
                    End If
                        
                Next lColCount
    
            Next lRowCount
            
        End If
        
        'Last Sub Total
        '--------------------------------------------------------------------------
        If .MaxRows > 0 Then
        
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1

            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HFFE6E6)
            
            For k = 1 To iDupCnt
                .Col = k
                If .ColHidden = False Then
                    .Text = "小  计"
                    
                    k = iDupCnt
                End If
            Next k
            
            For k = 1 To iSumCnt
                .Col = iSumCol(k)

                If iSumCol(k) <= 26 Then
                    sCol_a = Chr(iSumCol(k) + 64)
                    .Formula = "sum(" + sCol_a & iOld_Row & ":" + sCol_a & .Row - 1 & ")"
                Else
                    iCot = Int(((iSumCol(k) - 1) / 26))
                    iBas = 26 * iCot

                    sCol_a = Chr((iSumCol(k) - iBas) + 64)
                    sCol_b = Chr(iCot + 64)

                    .Formula = "sum(" + sCol_b + sCol_a & iOld_Row & ":" + sCol_b + sCol_a & .Row - 1 & ")"
                End If
                
                If VarType(.Value) = vbNull Then
                    dTotal(k) = dTotal(k) + 0
                Else
                    dTotal(k) = dTotal(k) + .Value
                End If
                
            Next k


        'Total Compute
        '--------------------------------------------------------------------------
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1

            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
            
            For k = 1 To iDupCnt
                .Col = k
                If .ColHidden = False Then
                    .Text = "合  计"
                    k = iDupCnt
                End If
            Next k
            
            For k = 1 To iSumCnt
                .Col = iSumCol(k)
                .Value = dTotal(k)
            Next k
            
        End If

        .ReDraw = True
        
        Gf_Stotal_Display = True
        Sc.Item("Spread").OperationMode = OperationModeRow
   
        Screen.MousePointer = vbDefault
    End With
    
    Exit Function

Error_Rtn:

    Set AdoRs = Nothing
    Gf_Stotal_Display = False

    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Multi_Stotal_Display
'   2.Name         : Muti Sub, Total Display
'   3.Input  Value : Conn Connection, Sc Collection, sQuery String,
'                    {iDupCnt Variant}, {iSumCnt Variant}, {iSumCol Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Sub, Total Display
'---------------------------------------------------------------------------------------
Public Function Gf_Multi_Stotal_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional iFrNo As Variant = 0, Optional iToNo As Variant = 0, _
                                  Optional iSumCnt As Variant = 0, Optional iSumCol As Variant, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn

    Dim I, j, k, l  As Long

    Dim iBas        As Integer
    Dim iCot        As Integer
    Dim iOld_Row    As Integer
    Dim Sw          As Boolean

    Dim lRowCount   As Long
    Dim lColCount   As Long

    Dim sCol_a      As String
    Dim sCol_b      As String
    Dim sTemp()     As String
    Dim sTitle()    As String

    Dim dSubTotal() As Double
    Dim dTotal()    As Double

    Dim AdoRs       As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check      ...............
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Multi_Stotal_Display = False: Exit Function
    End If

    If iSumCnt <> iSumCol.Count Then
        Call Gp_MsgBoxDisplay("iSumCnt and iSumCol are different..!!", "I")
        Gf_Multi_Stotal_Display = False
        Exit Function
    End If

    Set AdoRs = New ADODB.Recordset

    ReDim sTemp(iFrNo - 1 To iToNo - 1) As String
    ReDim sTitle(iFrNo To iToNo) As String
    ReDim dSubTotal(iFrNo To iToNo, 1 To iSumCnt) As Double
    ReDim dTotal(1 To iSumCnt) As Double

    With Sc.Item("Spread")

        Gf_Multi_Stotal_Display = True

        .ReDraw = False
        .MaxRows = 0
        iOld_Row = 1
        
'        For k = iFrNo To iToNo
'            For j = 1 To iSumCnt
'                dSubTotal(k, j) = 0
'            Next j
'            .Row = 0
'            .Col = k
'            sTitle(k) = .Text & "小计"
'        Next k

        Screen.MousePointer = vbHourglass

        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset

        If AdoRs.BOF Or AdoRs.EOF Then

            Call Gp_MsgBoxDisplay("无相关记录", "I")

            Gf_Multi_Stotal_Display = False
            .ReDraw = True

            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function

        End If

        ArrayRecords = AdoRs.GetRows

        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 1) <> 0 Then

            For lRowCount = 0 To UBound(ArrayRecords, 2)

                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                Sw = False

                'Duplicate Process
                For j = iFrNo To iToNo Step 1

                    If sTemp(j - 1) <> Trim(ArrayRecords(j - 1, lRowCount)) Then

                        'Sub Total
                        '--------------------------------------------------------------------------
                        If .Row <> 1 Then 'And j = 1 Then

                            For I = iToNo To j Step -1
                                .Col = 1
                                Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .Row, .Row, BLACK, &HFFE6E6)
                                
                                .Col = I
                                If .ColHidden = False Then
'                                    .Text = sTitle(I)
                                    .Text = sTemp(I - 1) & " 小计"
                                End If
                                
                                If I = iToNo Then
                                    For k = 1 To iSumCnt
                                        .Col = iSumCol(k)
    
                                        If iSumCol(k) <= 26 Then
                                            sCol_a = Chr(iSumCol(k) + 64)
                                            .Formula = "sum(" + sCol_a & iOld_Row & ":" + sCol_a & .Row - 1 & ")"
                                        Else
                                            iCot = Int(((iSumCol(k) - 1) / 26))
                                            iBas = 26 * iCot
    
                                            sCol_a = Chr((iSumCol(k) - iBas) + 64)
                                            sCol_b = Chr(iCot + 64)
    
                                            .Formula = "sum(" + sCol_b + sCol_a & iOld_Row & ":" + sCol_b + sCol_a & .Row - 1 & ")"
                                        End If
                                        
                                        For l = iFrNo To I
                                            dSubTotal(l, k) = dSubTotal(l, k) + Val(.Value & "")
                                        Next l
    
                                        If VarType(.Value) = vbNull Then
                                            dTotal(k) = dTotal(k) + 0
                                        Else
                                            dTotal(k) = dTotal(k) + Val(.Value & "")
                                        End If
                                    Next k
                                Else
                                    For k = 1 To iSumCnt
                                        .Col = iSumCol(k)
                                        .Value = dSubTotal(I, k)
                                        dSubTotal(I, k) = 0
                                    Next k
                                End If
                                
                                .MaxRows = .MaxRows + 1
                                .Row = .MaxRows
                                iOld_Row = .Row

                            Next I
                            
                            For k = iToNo To j - 1 Step -1
                                If k = 0 Then Exit For
                                sTemp(k - 1) = ArrayRecords(k - 1, lRowCount)
                            Next k

                            Sw = True
                            Exit For

                        End If
                        '--------------------------------------------------------------------------
                        .Col = j
                        If VarType(ArrayRecords(j - 1, lRowCount)) = 8 Then
                            .Text = ArrayRecords(j - 1, lRowCount)
                        Else
                            .Value = Val(ArrayRecords(j - 1, lRowCount) & "")
                        End If

                        sTemp(j - 1) = ArrayRecords(j - 1, lRowCount)

                        For k = j + 1 To iToNo Step 1
                            sTemp(k - 1) = ""
                        Next k

                    End If

                Next j

                'Duplicate Process

                If Sw And j > iToNo Then
                    For lColCount = 0 To .MaxCols - 1
                        .Col = lColCount + 1

                        If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(lColCount, lRowCount))
                        End If

                    Next lColCount
                Else
                    If MsgChk Then
                        If iFrNo > 1 Then
                            For lColCount = 0 To iFrNo - 2
                                .Col = lColCount + 1
    
                                If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                                    .Text = ""
                                Else
                                    .Text = Trim(ArrayRecords(lColCount, lRowCount))
                                End If
    
                            Next lColCount
                        End If
    
                        For lColCount = j - 1 To .MaxCols - 1
                            .Col = lColCount + 1
    
                            If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(lColCount, lRowCount))
                            End If
    
                        Next lColCount
                    Else
                        For lColCount = 0 To .MaxCols - 1
                            .Col = lColCount + 1
    
                            If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(lColCount, lRowCount))
                            End If
    
                        Next lColCount
                    End If
                    
                End If
            Next lRowCount

        End If

        'Last Sub Total
        '--------------------------------------------------------------------------
        If .MaxRows > 0 Then

            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            For I = iToNo To iFrNo Step -1
            
                .Col = 1
                Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HFFE6E6)
                
                .Col = I
                If .ColHidden = False Then
'                    .Text = sTitle(I)
                    .Text = sTemp(I - 1) & " 小计"
                End If
                
                If I = iToNo Then
                    For k = 1 To iSumCnt
                        .Col = iSumCol(k)

                        If iSumCol(k) <= 26 Then
                            sCol_a = Chr(iSumCol(k) + 64)
                            .Formula = "sum(" + sCol_a & iOld_Row & ":" + sCol_a & .Row - 1 & ")"
                        Else
                            iCot = Int(((iSumCol(k) - 1) / 26))
                            iBas = 26 * iCot

                            sCol_a = Chr((iSumCol(k) - iBas) + 64)
                            sCol_b = Chr(iCot + 64)

                            .Formula = "sum(" + sCol_b + sCol_a & iOld_Row & ":" + sCol_b + sCol_a & .Row - 1 & ")"
                        End If
                        
                        For l = iFrNo To I
                            dSubTotal(l, k) = dSubTotal(l, k) + Val(.Value & "")
                        Next l

                        If VarType(.Value) = vbNull Then
                            dTotal(k) = dTotal(k) + 0
                        Else
                            dTotal(k) = dTotal(k) + Val(.Value & "")
                        End If
                    Next k
                Else
                    For k = 1 To iSumCnt
                        .Col = iSumCol(k)
                        .Value = dSubTotal(I, k)
                        dSubTotal(I, k) = 0
                    Next k
                End If
                
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            Next I


        'Total Compute
'        '--------------------------------------------------------------------------
'            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1

            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)

            For k = iFrNo To iToNo
                .Col = k
                If .ColHidden = False Then
                    .Text = "合计"
                    k = iToNo
                End If
            Next k

            For k = 1 To iSumCnt
                .Col = iSumCol(k)
                .Value = dTotal(k)
            Next k

        End If

        .ReDraw = True

        Gf_Multi_Stotal_Display = True
        Sc.Item("Spread").OperationMode = OperationModeRow

        Screen.MousePointer = vbDefault
    End With

    Exit Function

Error_Rtn:

    Set AdoRs = Nothing
    Gf_Multi_Stotal_Display = False

    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sub_total_Display
'   2.Name         : Sub, Total Display
'   3.Input  Value : Conn Connection, Sc Collection, sQuery String,
'                    {iDupCnt Variant}, {iSumCnt Variant}, {iSumCol Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : ZHANG.LIN
'   6.Create Date  : 2006. 10 .18
'   7.Modify Date  :
'   8.Comment      : Sub, Total Display
'---------------------------------------------------------------------------------------
Public Function Gf_Sub_total_Display(Conn As ADODB.Connection, Sc As Collection, sQuery As String, Optional iDupCnt As Variant = 0, _
                                  Optional iSumCnt As Variant = 0, Optional iSumCol As Variant, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Error_Rtn

    Dim iDX_K, iDX_J, iDX_M, iDX_N, iDX_L As Integer
    
    Dim iBas As Integer
    Dim iCot As Integer
    Dim read_cnt As Integer
    Dim iOld_Row As Integer
    
    Dim Sw As Boolean

    Dim lRowCount As Long
    Dim lColCount As Long

    Dim sCol_a As String
    Dim sCol_b As String
    Dim sTemp() As String
    Dim vTemp As String

    Dim dTotal() As Double
    Dim dSUB_Total() As Double

    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Sub_total_Display = False: Exit Function
    End If
    
    If iSumCnt <> iSumCol.Count Then
        Call Gp_MsgBoxDisplay("iSumCnt and iSumCol are different..!!", "I")
        Gf_Sub_total_Display = False
        Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    ReDim sTemp(0 To iDupCnt - 1) As String
    ReDim dTotal(1 To iSumCnt) As Double
    ReDim dSUB_Total(1 To iSumCnt) As Double

    For iDX_K = 1 To iSumCnt
        dSUB_Total(iDX_K) = 0
        dTotal(iDX_K) = 0
    Next iDX_K
        
    With Sc.Item("Spread")
    
        Gf_Sub_total_Display = True
        
        .ReDraw = False
        .MaxRows = 0
        iOld_Row = 1
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
        
            Gf_Sub_total_Display = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
        
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 1) <> 0 Then
        
            For lRowCount = 0 To UBound(ArrayRecords, 2)
            
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                'Duplicate Process
                For iDX_J = 1 To iDupCnt Step 1
                
                    If lRowCount = 0 Then
                       For iDX_M = 1 To iDupCnt Step 1
                           sTemp(iDX_M - 1) = Trim(ArrayRecords(iDX_M - 1, lRowCount))
                       Next iDX_M
                    End If
                
                    If sTemp(iDX_J - 1) <> Trim(ArrayRecords(iDX_J - 1, lRowCount)) Then
    
                        'Sub Total
                        '--------------------------------------------------------------------------
                        If iDX_J = 1 Then
                        
                            .Col = 1
    
                            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .Row, .Row, BLACK, &HFFE6E6)
                            
                            For iDX_N = 1 To iDupCnt
                                .Col = iDX_N
                                If .ColHidden = False Then
                                    .Text = sTemp(iDX_J - 1) & "小  计"
                                    iDX_N = iDupCnt
                                End If
                            Next iDX_N
                            
                            For iDX_L = 1 To iSumCnt
                                .Col = iSumCol(iDX_L)
                                .Value = dSUB_Total(iDX_L)
                            Next iDX_L
                            
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            iOld_Row = .Row

                            For iDX_K = 1 To iSumCnt
                                    dSUB_Total(iDX_K) = 0
                            Next iDX_K
                            
                        End If
                        '--------------------------------------------------------------------------
                        .Col = iDX_J
                        .Text = ArrayRecords(iDX_J - 1, lRowCount)
                      
                    End If
                    
                    '--------------------------------------------------------------------------
                    .Col = iDX_J
                    
                    If lRowCount > 0 And _
                       (VarType(ArrayRecords(iDX_J - 1, lRowCount)) = vbNull Or _
                        sTemp(iDX_J - 1) = ArrayRecords(iDX_J - 1, lRowCount)) Then
                       .Text = ""
                    Else
                       .Text = ArrayRecords(iDX_J - 1, lRowCount)
                    End If
                    
                    If VarType(ArrayRecords(iDX_J - 1, lRowCount)) = vbNull Then
                        sTemp(iDX_J - 1) = ""
                    Else
                        sTemp(iDX_J - 1) = ArrayRecords(iDX_J - 1, lRowCount)
                    End If
                    
'                    For idx_l = idx_J + 1 To iDupCnt Step 1
'                        sTemp(idx_l - 1) = ""
'                    Next idx_l
                    
                    If iDX_J = 1 Then
                        For iDX_K = 1 To iSumCnt
                            
                            If VarType(ArrayRecords(iSumCol(iDX_K) - 1, lRowCount)) = vbNull Then
                                dTotal(iDX_K) = dTotal(iDX_K) + 0
                                dSUB_Total(iDX_K) = dSUB_Total(iDX_K) + 0
                            Else
                                dTotal(iDX_K) = dTotal(iDX_K) + ArrayRecords(iSumCol(iDX_K) - 1, lRowCount)
                                dSUB_Total(iDX_K) = dSUB_Total(iDX_K) + ArrayRecords(iSumCol(iDX_K) - 1, lRowCount)
                            End If
                            
                        Next iDX_K
                    End If
                                            
                Next iDX_J
                
                'Duplicate Process
                For lColCount = iDupCnt To .MaxCols - 1
                    .Col = lColCount + 1
                    
                    If VarType(ArrayRecords(lColCount, lRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(lColCount, lRowCount))
                    End If
                        
                Next lColCount
    
            Next lRowCount
            
        End If
        
        'Last Sub Total
        '--------------------------------------------------------------------------
        If .MaxRows > 0 Then
        
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1

            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HFFE6E6)
            
            For iDX_M = 1 To iDupCnt
                .Col = iDX_M
                If .ColHidden = False Then
                    .Text = sTemp(iDX_M - 1) & "小  计"
                    iDX_M = iDupCnt
                End If
            Next iDX_M
            
            For iDX_L = iDX_J + 1 To iDupCnt Step 1
                sTemp(iDX_L - 1) = ""
            Next iDX_L
            
            For iDX_N = 1 To iSumCnt
                .Col = iSumCol(iDX_N)
                .Value = dSUB_Total(iDX_N)
            Next iDX_N

        'Total Compute
        '--------------------------------------------------------------------------
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1

            Call Gp_Sp_BlockColor(Sc.Item("Spread"), 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
            
            For iDX_M = 1 To iDupCnt
                .Col = iDX_M
'                If .ColHidden = False Then
                    .Text = "合  计"
                    iDX_M = iDupCnt
'                End If
            Next iDX_M
            
            For iDX_L = 1 To iSumCnt
                .Col = iSumCol(iDX_L)
                .Value = dTotal(iDX_L)
            Next iDX_L
            
        End If

        .ReDraw = True
        
        Gf_Sub_total_Display = True
        Sc.Item("Spread").OperationMode = OperationModeRow
   
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

Error_Rtn:

'    Set adoRs = Nothing
    Gf_Sub_total_Display = False

    Screen.MousePointer = vbDefault

End Function

