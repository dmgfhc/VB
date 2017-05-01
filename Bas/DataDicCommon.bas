Attribute VB_Name = "DataDicCommon"
Option Explicit

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_DD_Display
'   2.Name         : Data Dictionary Result Display
'   3.Input  Value : Conn Connection, sQuery String, [MsgChk Boolean]
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  : 2003. 07 .31
'   8.Comment      : Data Dictionary Result Data Display
'---------------------------------------------------------------------------------------
Public Function Gf_DD_Display(Conn As ADODB.Connection, sQuery As String, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo DD_Display_Error

    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then
            Gf_DD_Display = False
            DD.DataDicType = ""
            DD.DicRefType = ""
            DD.nameType = ""
            DD.sQuery = ""
            DD.sSelect = False
            DD.sWitch = ""
            DD.sWhere = ""
            DD.sKey = ""
            
            Set DD.rControl = Nothing
            Set DD.wControl = Nothing
            Set DD.sPname = Nothing
            Exit Function
        End If
    End If
    
    Load DataDic
    
    Set AdoRs = New ADODB.Recordset
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
        
    With DataDic

        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
            
            .ssResult.MaxCols = AdoRs.Fields.Count
            .ssWhere.MaxCols = AdoRs.Fields.Count
            
            .ssResult.MaxRows = 0
            .ssWhere.MaxRows = 1
            
            .ssResult.ROW = 0
            .ssWhere.ROW = 0
            
            .ssResult.Col = -1
            .ssResult.ROW = 0
            .ssResult.FontBold = True
            
            .ssWhere.Col = -1
            .ssWhere.ROW = 0
            .ssWhere.FontBold = True
            
            For iColcount = 0 To .ssResult.MaxCols - 1
                    
                .ssResult.Col = iColcount + 1
                .ssWhere.Col = iColcount + 1
                
                If VarType(AdoRs.Fields(iColcount).Name) = vbNull Then
                    .ssResult.Text = ""
                    .ssWhere.Text = ""
                Else
                    .ssResult.Text = Trim(AdoRs.Fields(iColcount).Name)
                    .ssWhere.Text = Trim(AdoRs.Fields(iColcount).Name)
                End If
                
                .ssResult.TypeMaxEditLen = AdoRs.Fields(iColcount).DefinedSize
                .ssWhere.TypeMaxEditLen = AdoRs.Fields(iColcount).DefinedSize
                        
            Next iColcount
                
            If DD.DicRefType = "C" Then DataDic.Show 1
            
            Gf_DD_Display = False
            AdoRs.Close
            Set AdoRs = Nothing
            Exit Function
            
        End If
    
        Screen.MousePointer = vbHourglass
        
        Gf_DD_Display = True
        
        .ssResult.ReDraw = False
        .ssWhere.ReDraw = False
        
        .ssResult.MaxRows = 0
        .ssWhere.MaxRows = 1
        
        'Result Spread Column Name Setting
        .ssResult.ROW = 0
        .ssWhere.ROW = 0
        
        .ssResult.MaxCols = AdoRs.Fields.Count
        .ssWhere.MaxCols = AdoRs.Fields.Count
        
        .ssResult.Col = -1
        .ssResult.ROW = 0
        .ssResult.FontBold = True
        
        .ssWhere.Col = -1
        .ssWhere.ROW = 0
        .ssWhere.FontBold = True
        
        For iColcount = 0 To .ssResult.MaxCols - 1
                
            .ssResult.Col = iColcount + 1
            .ssWhere.Col = iColcount + 1
            
            If VarType(AdoRs.Fields(iColcount).Name) = vbNull Then
                .ssResult.Text = ""
                .ssWhere.Text = ""
            Else
                .ssResult.Text = Trim(AdoRs.Fields(iColcount).Name)
                .ssWhere.Text = Trim(AdoRs.Fields(iColcount).Name)
            End If
            
            .ssResult.TypeMaxEditLen = AdoRs.Fields(iColcount).DefinedSize
            .ssWhere.TypeMaxEditLen = AdoRs.Fields(iColcount).DefinedSize
                    
        Next iColcount
                
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 1) >= 0 Then
        
            .ssResult.MaxRows = UBound(ArrayRecords, 2) + 1
        
            For iRowCount = 0 To .ssResult.MaxRows - 1
            
                .ssResult.ROW = iRowCount + 1
                
                For iColcount = 0 To .ssResult.MaxCols - 1
                
                    .ssResult.Col = iColcount + 1
                    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                        .ssResult.Text = ""
                    Else
                        .ssResult.Text = Trim(ArrayRecords(iColcount, iRowCount))
                    End If
                            
                Next iColcount
                
            Next iRowCount
            
        End If
        
        .ssResult.ReDraw = True
        .ssWhere.ReDraw = True
        
        Screen.MousePointer = vbDefault
        
    End With

    If DD.DicRefType = "C" Then DataDic.Show 1
            
    Exit Function

DD_Display_Error:
    
    Unload DataDic
    
    Gf_DD_Display = False
    DD.DataDicType = ""
    DD.DicRefType = ""
    DD.nameType = ""
    DD.sQuery = ""
    DD.sSelect = False
    DD.sWitch = ""
    DD.sWhere = ""
    DD.sKey = ""
    
    Set DD.rControl = Nothing
    Set DD.wControl = Nothing
    Set DD.sPname = Nothing
    
    Set AdoRs = Nothing
    Gf_DD_Display = False
    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Common_DD
'   2.Name         : Common Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Common Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

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
        DD.sQuery = DD.sQuery + " WHERE CD_MANA_NO =    '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND CD         like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
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
        
        DD.sWhere = DD.sWhere + " ORDER  BY  CD  ASC "
    
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT CD ""代码"", CD_SHORT_NAME ""代码简称"", CD_NAME ""代码名称"", "
        DD.sQuery = DD.sQuery + "       CD_SHORT_ENG ""代码英文简称"", CD_FULL_ENG ""代码英文名称"" FROM NISCO.ZP_CD "
        DD.sQuery = DD.sQuery + " WHERE CD_MANA_NO =    '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND CD         like '" & Trim(DD.sPname.Text) & "%' "
        
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
            
            DD.sWhere = DD.sWhere + " ORDER  BY  CD  ASC "
            
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
'   1.ID           : Gf_Usage_DD
'   2.Name         : Order Usage Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Order Usage Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Usage_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
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

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
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
    
    DD.DataDicType = "U"        'Order Usage Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT ENDUSE_CD ""订单用途"", ENDUSE_NAME ""订单用途名称"" FROM NISCO.QP_ORD_USAGE "
        DD.sQuery = DD.sQuery + " WHERE PROD_KND             LIKE   '" & Trim(DD.sKey) & "%' "
        DD.sWhere = DD.sWhere + "   AND ENDUSE_CD            like   '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(ENDUSE_NAME,'%') like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  ENDUSE_CD  ASC "
        
    Else
    
        DD.sQuery = "            SELECT ENDUSE_CD ""订单用途"", ENDUSE_NAME ""订单用途名称"" FROM NISCO.QP_ORD_USAGE "
        DD.sQuery = DD.sQuery + " WHERE PROD_KND = '" & Trim(DD.sKey) & "' "
        
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sWhere = DD.sWhere + "     AND ENDUSE_CD            like '" & Trim(DD.sPname.Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            
            DD.sWhere = DD.sWhere + " AND NVL(ENDUSE_NAME,'%') like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  ENDUSE_CD  ASC "

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
'   1.ID           : Gf_Customer_DD
'   2.Name         : Customer Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Customer Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Customer_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
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
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "C"        'Customer Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT  CUST_CD ""客户代码"", CUST_NM ""客户名称"", CUST_NM_ENG ""客户英文名称"" FROM NISCO.BP_CUST_CD "
        DD.sWhere = " WHERE  CUST_CD           like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM,'%')      like '" & Trim(DD.rControl.Item(2).Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM_ENG,'%')  like '" & Trim(DD.rControl.Item(2).Text) & "%' "
            End Select
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  CUST_CD  ASC "
    
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = " SELECT  CUST_CD ""客户代码"", CUST_NM ""客户名称"", CUST_NM_ENG ""客户英文名称"" FROM NISCO.BP_CUST_CD "
        DD.sWhere = "  WHERE  CUST_CD           like '" & Trim(DD.sPname.Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            
            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM,'%')      like '" & Trim(DD.sPname.Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM_ENG,'%')  like '" & Trim(DD.sPname.Text) & "%' "
            End Select
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  CUST_CD  ASC "
   
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
'   1.ID           : Gf_Customer_DD
'   2.Name         : Customer Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Customer Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Customer_DD2(Conn As ADODB.Connection, KeyCode As Integer, BySalComFl As String) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
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
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "C"        'Customer Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT  CUST_CD ""客户代码"", CUST_NM ""客户名称"", CUST_NM_ENG ""客户英文名称"" FROM NISCO.BP_CUST_CD "
        DD.sWhere = " WHERE  CUST_CD           like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND  CUST_TYP  IN ( 'Z','" & BySalComFl & "')"
        If DD.rControl.Count > 1 Then
            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM,'%')      like '" & Trim(DD.rControl.Item(2).Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM_ENG,'%')  like '" & Trim(DD.rControl.Item(2).Text) & "%' "
            End Select
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  CUST_CD  ASC "
    
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = " SELECT  CUST_CD ""客户代码"", CUST_NM ""客户名称"", CUST_NM_ENG ""客户英文名称"" FROM NISCO.BP_CUST_CD "
        DD.sWhere = "  WHERE  CUST_CD           like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND  CUST_TYP  IN ( 'Z','" & BySalComFl & "')"
        
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            
            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM,'%')      like '" & Trim(DD.sPname.Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(CUST_NM_ENG,'%')  like '" & Trim(DD.sPname.Text) & "%' "
            End Select
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  CUST_CD  ASC "
   
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
'   1.ID           : Gf_Destination_DD
'   2.Name         : Destination Code Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Destination Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Destination_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "D"        'Destination Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT DEST_CD ""目的地代码"", CITY_CD ""城市代码"", STATION_CD ""站台代码"", DEST_NM ""目的地名"", "
        DD.sQuery = DD.sQuery + "       DEST_NM_ENG ""目的地英文名"", DEST_ADDR ""经营地址"", POST ""邮编号"", DOME_FL ""订单分类名称"", "
        DD.sQuery = DD.sQuery + "       COUNTRY_CD ""国家代码"" FROM  NISCO.BP_DEST_CD "
        DD.sWhere = "             WHERE DEST_CD                          like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(DEST_NM,'%')       like '" & Trim(DD.rControl.Item(2).Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(DEST_NM_ENG,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
            End Select
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  DEST_CD  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT DEST_CD ""目的地代码"", CITY_CD ""城市代码"", STATION_CD ""站台代码"", DEST_NM ""目的地名"", "
        DD.sQuery = DD.sQuery + "       DEST_NM_ENG ""目的地英文名"", DEST_ADDR ""经营地址"", POST ""邮编号"", DOME_FL ""订单分类名称"", "
        DD.sQuery = DD.sQuery + "       COUNTRY_CD ""国家代码"" FROM  NISCO.BP_DEST_CD "
        DD.sWhere = "             WHERE DEST_CD                          like '" & Trim(DD.sPname.Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            
            Select Case DD.nameType
                Case "1"
                    DD.sWhere = DD.sWhere + " AND NVL(DEST_NM,'%')       like '" & Trim(DD.sPname.Text) & "%' "
                Case "2"
                    DD.sWhere = DD.sWhere + " AND NVL(DEST_NM_ENG,'%')   like '" & Trim(DD.sPname.Text) & "%' "
            End Select
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  DEST_CD  ASC "
        
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
'   1.ID           : Gf_Apply_DD
'   2.Name         : Apply Item Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .19
'   7.Modify Date  :
'   8.Comment      : Apply Item Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Apply_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "A"        'Apply Item Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT APLY_ITEM ""适用项目"", APLY_ITEM_NAME ""适用项目名称"" FROM  NISCO.ZP_APLY_ITEM "
        DD.sQuery = DD.sQuery + " WHERE TABLE_ID                  =    '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND APLY_ITEM                 like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(APLY_ITEM_NAME,'%') like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  APLY_ITEM  ASC "
        
    Else
    
        DD.sQuery = "            SELECT APLY_ITEM ""适用项目"", APLY_ITEM_NAME ""适用项目名称"" FROM  NISCO.ZP_APLY_ITEM "
        DD.sQuery = DD.sQuery + " WHERE TABLE_ID                  = '" & Trim(DD.sKey) & "' "
        
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sWhere = DD.sWhere + " AND APLY_ITEM                   like '" & Trim(DD.sPname.Text) & "%' "
        
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            
            DD.sWhere = DD.sWhere + " AND NVL(APLY_ITEM_NAME,'%') like '" & Trim(DD.sPname.Text) & "%' "
        End If

        DD.sWhere = DD.sWhere + " ORDER  BY  APLY_ITEM  ASC "
        
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
'   1.ID           : Gf_Stlgrd_DD
'   2.Name         : Stlgrd Code Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .20
'   7.Modify Date  :
'   8.Comment      : Stlgrd Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Stlgrd_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "S"        'Stlgrd Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.rControl.Item(1).Text) & "%' "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.sPname.Text) & "%' "
            
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
   
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
'   1.ID           : Gf_StdSPEC_DD
'   2.Name         : StdSPEC Code Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .20
'   7.Modify Date  :
'   8.Comment      : StdSPEC Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_StdSPEC_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "T"        'StdSPEC Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.rControl.Item(1).Text) & "%' "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.sPname.Text) & "%' "
            
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
   
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
'   1.ID           : Gf_StdSPEC_DD2
'   2.Name         : StdSPEC Code Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .20
'   7.Modify Date  :
'   8.Comment      : StdSPEC Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_StdSPEC_DD2(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "T"        'StdSPEC Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.rControl.Item(1).Text) & "%'  AND NVL(STDSPEC_CHR_CD,'Y') <>'N' "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.sPname.Text) & "%'  AND NVL(STDSPEC_CHR_CD,'Y') <>'N' "
            
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
   
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
'   1.ID           : Gf_Nisco_STD_DD
'   2.Name         : 企标材质 - 企业标准编号
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 08. 04
'   7.Modify Date  :
'   8.Comment      : 企标材质 - 企业标准编号
'---------------------------------------------------------------------------------------
Public Function Gf_Nisco_STD_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "N"   'Nisco Standard No
    DD.DicRefType = "C"    'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT NISCO_QUALITY_NO ""企业标准编号"",APP_Date ""下达日期"","
        DD.sQuery = DD.sQuery + "       (Trim(TO_CHAR(THK_MIN,'000.00'))||'~'||Trim(TO_CHAR(THK_MAX,'000.00'))) ""厚度分组"" FROM  NISCO.QP_NISCO_MATR "
        DD.sWhere = "             WHERE NISCO_QUALITY_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(APP_DATE,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  NISCO_QUALITY_NO  ASC "
            
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT NISCO_QUALITY_NO ""企业标准编号"",APP_Date ""下达日期"","
        DD.sQuery = DD.sQuery + "       (Trim(TO_CHAR(THK_MIN,'000.00'))||'~'||Trim(TO_CHAR(THK_MAX,'000.00'))) ""厚度分组"" FROM  NISCO.QP_NISCO_MATR "
        DD.sWhere = "             WHERE NISCO_QUALITY_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
                    
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(APP_DATE,'%')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  NISCO_QUALITY_NO  ASC "
        
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
'   1.ID           : Gf_Roll_STD_DD
'   2.Name         : 轧钢生产规范 - 轧钢生产规范编号
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 08. 04
'   7.Modify Date  :
'   8.Comment      : 轧钢生产规范 - 轧钢生产规范编号
'---------------------------------------------------------------------------------------
Public Function Gf_Roll_STD_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 1 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "R"   'MILL_STD_NO
    DD.DicRefType = "C"    'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT MILL_STD_NO ""轧钢生产规范编号"", APP_DATE ""下达日期"","
        DD.sQuery = DD.sQuery + "       ('MIN: '||TO_CHAR(THK_MIN,'000.00')||' MAX: '||TO_CHAR(THK_MAX,'000.00')) ""厚度分组"", GF_STLGRD_DETAIL(STLGRD) ""钢种"" FROM  NISCO.QP_ROLL_STD "
        DD.sWhere = "             WHERE MILL_STD_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  MILL_STD_NO  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT MILL_STD_NO ""轧钢生产规范编号"", APP_DATE ""下达日期"","
        DD.sQuery = DD.sQuery + "       ('MIN: '||TO_CHAR(THK_MIN,'000.00')||' MAX: '||TO_CHAR(THK_MAX,'000.00')) ""厚度分组"", GF_STLGRD_DETAIL(STLGRD) ""钢种"" FROM  NISCO.QP_ROLL_STD "
        DD.sWhere = "             WHERE MILL_STD_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  MILL_STD_NO  ASC "
        
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
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
'   1.ID           : Gf_MILL_STD_DD
'   2.Name         : 轧钢生产规范 - 轧钢生产规范编号
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 08. 04
'   7.Modify Date  :
'   8.Comment      : 轧钢生产规范 - 轧钢生产规范编号
'---------------------------------------------------------------------------------------
Public Function Gf_MILL_STD_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 1 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "R"   'MILL_STD_NO
    DD.DicRefType = "C"    'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT MILL_STD_NO ""轧钢生产规范编号"", APP_DATE ""下达日期"","
        DD.sQuery = DD.sQuery + "      (TO_CHAR(THK_MIN,'000.00')||' ~ '||TO_CHAR(THK_MAX,'000.00')) ""厚度分组"", "
        DD.sQuery = DD.sQuery + "      (TO_CHAR(WID_MIN,'0000.00')||' ~ '||TO_CHAR(WID_MAX,'0000.00')) ""贯度分组"", "
        DD.sQuery = DD.sQuery + "       GF_STLGRD_DETAIL(STLGRD) ""钢种"" FROM  NISCO.QP_MILL_STD "
        DD.sWhere = "             WHERE MILL_STD_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  MILL_STD_NO  ASC "
            
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT MILL_STD_NO ""轧钢生产规范编号"", APP_DATE ""下达日期"","
        DD.sQuery = DD.sQuery + "      (TO_CHAR(THK_MIN,'000.00')||' ~ '||TO_CHAR(THK_MAX,'000.00')) ""厚度分组"", "
        DD.sQuery = DD.sQuery + "      (TO_CHAR(WID_MIN,'0000.00')||' ~ '||TO_CHAR(WID_MAX,'0000.00')) ""贯度分组"", "
        DD.sQuery = DD.sQuery + "       GF_STLGRD_DETAIL(STLGRD) ""钢种"" FROM  NISCO.QP_MILL_STD "
        DD.sWhere = "             WHERE MILL_STD_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  MILL_STD_NO  ASC "
  
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
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
'   1.ID           : Gf_Melt_STD_DD
'   2.Name         : 炼钢/连铸生产规范 - 炼钢/连铸生产规范编号
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 08. 04
'   7.Modify Date  :
'   8.Comment      : 炼钢/连铸生产规范 - 炼钢/连铸生产规范编号
'---------------------------------------------------------------------------------------
Public Function Gf_Melt_STD_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "L"   'MLT_STD_NO
    DD.DicRefType = "C"    'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT MLT_STD_NO ""炼钢/连铸生产规范编号"",APP_DATE ""下达日期"", GF_STLGRD_DETAIL(STLGRD) ""钢种"" FROM  NISCO.QP_MELT_STD "
        DD.sWhere = " WHERE MLT_STD_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  MILL_STD_NO  ASC "
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "SELECT MLT_STD_NO ""炼钢/连铸生产规范编号"",APP_DATE ""下达日期"", GF_STLGRD_DETAIL(STLGRD) ""钢种"" FROM  NISCO.QP_MELT_STD "
        DD.sWhere = " WHERE MLT_STD_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  MILL_STD_NO  ASC "
   
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
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
'   1.ID           : Gf_Cust_STD_DD
'   2.Name         : 客户特殊要求共用信息 - 客户特殊要求编号
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 08. 04
'   7.Modify Date  :
'   8.Comment      : 客户特殊要求共用信息 - 客户特殊要求编号
'---------------------------------------------------------------------------------------
Public Function Gf_Cust_STD_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "E"   'CUST_SPEC_NO
    DD.DicRefType = "C"    'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT CUST_SPEC_NO ""客户特殊要求编号"", GF_CUSTNAMEFIND(SUBSTR(CUST_SPEC_NO,1,6)) ""客户名称"", "
        DD.sQuery = DD.sQuery + "       PROD_CD ""产品代码"" FROM  NISCO.QP_CUST_HEAD "
        DD.sWhere = "             WHERE CUST_SPEC_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  CUST_SPEC_NO  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT CUST_SPEC_NO ""客户特殊要求编号"", GF_CUSTNAMEFIND(SUBSTR(CUST_SPEC_NO,1,6)) ""客户名称"", "
        DD.sQuery = DD.sQuery + "       PROD_CD ""产品代码"" FROM  NISCO.QP_CUST_HEAD "
        DD.sWhere = "             WHERE CUST_SPEC_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  CUST_SPEC_NO  ASC "
            
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
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
'   1.ID           : Gf_STD_DELV_DD
'   2.Name         : 标准交付条件 - 代表性交付条件标准编号
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Lee Qing Yu
'   6.Create Date  : 2003. 08. 04
'   7.Modify Date  :
'   8.Comment      : 标准交付条件 - 代表性交付条件标准编号
'---------------------------------------------------------------------------------------
Public Function Gf_STD_DELV_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 1 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "V"  'DEV_STD_CD
    DD.DicRefType = "C"   'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT DISTINCT DEV_STD_CD ""代表性交付条件标准编号"" "
        DD.sQuery = DD.sQuery & "  FROM NISCO.QP_STD_DELV "
        DD.sWhere = "             WHERE DEV_STD_CD like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  DEV_STD_CD  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT DISTINCT DEV_STD_CD ""代表性交付条件标准编号"" "
        DD.sQuery = DD.sQuery & "  FROM NISCO.QP_STD_DELV "
        DD.sWhere = "             WHERE DEV_STD_CD like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  DEV_STD_CD  ASC "
   
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
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
'   1.ID           : Gf_ThkGrp_DD
'   2.Name         : THK GROUP Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 09 .26
'   7.Modify Date  :
'   8.Comment      : THK GROUP Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_ThkGrp_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

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
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 3 Or DD.nameType = "" Then
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
    
    DD.DataDicType = "G"        'THK GROUP Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT THK_CD ""厚度组"", FR_THK ""厚度下限"", TO_THK ""厚度上限"" FROM  NISCO.BP_THICK_GRP "
        DD.sQuery = DD.sQuery + " WHERE PROD_CD  = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND THK_CD   like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  THK_CD  ASC, FR_THK ASC,  TO_THK ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT THK_CD ""厚度组"", FR_THK ""厚度下限"", TO_THK ""厚度上限"" FROM  NISCO.BP_THICK_GRP "
        DD.sQuery = DD.sQuery + " WHERE PROD_CD  = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND THK_CD   like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  THK_CD  ASC, FR_THK ASC,  TO_THK ASC "
        
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
'   1.ID           : Gf_WidGrp_DD
'   2.Name         : WID GROUP Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 09 .26
'   7.Modify Date  :
'   8.Comment      : WID GROUP Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_WidGrp_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

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
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 3 Or DD.nameType = "" Then
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
    
    DD.DataDicType = "W"        'WID GROUP Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT WID_CD ""宽度组"", FR_WID ""宽度下限"", TO_WID ""宽度上限"" FROM  NISCO.BP_WIDTH_GRP "
        DD.sQuery = DD.sQuery + " WHERE PROD_CD  = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND WID_CD   like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  WID_CD  ASC, FR_WID ASC,  TO_WID ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT WID_CD ""宽度组"", FR_WID ""宽度下限"", TO_WID ""宽度上限"" FROM  NISCO.BP_WIDTH_GRP "
        DD.sQuery = DD.sQuery + " WHERE PROD_CD  = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND WID_CD   like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  WID_CD  ASC, FR_WID ASC,  TO_WID ASC "
        
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
'   1.ID           : Gf_Roll_ThkGrp_DD
'   2.Name         : ROLL THK GROUP Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 11 .6
'   7.Modify Date  :
'   8.Comment      : ROLL THK GROUP Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Roll_ThkGrp_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

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
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 3 Or DD.nameType = "" Then
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
    
    DD.DataDicType = "RTG"      'ROLL THK GROUP CODE
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT THK_GRP_CD ""厚度组"", MINI ""厚度下限"", MAXI ""厚度上限"" FROM  NISCO.EP_ROLL_THK_GRP "
        DD.sQuery = DD.sQuery + " WHERE PLT || PRC_LINE = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND THK_GRP_CD   like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  THK_GRP_CD  ASC, MINI ASC,  MAXI ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT THK_GRP_CD ""厚度组"", MINI ""厚度下限"", MAXI ""厚度上限"" FROM  NISCO.EP_ROLL_THK_GRP "
        DD.sQuery = DD.sQuery + " WHERE PLT || PRC_LINE = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND THK_GRP_CD   like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  THK_GRP_CD  ASC, MINI ASC,  MAXI ASC "
        
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
'   1.ID           : Gf_Roll_WidGrp_DD
'   2.Name         : ROLL WID GROUP Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 11 .6
'   7.Modify Date  :
'   8.Comment      : ROLL WID GROUP Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Roll_WidGrp_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

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
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 3 Or DD.nameType = "" Then
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
    
    DD.DataDicType = "RWG"      'ROLL WID GROUP CODE
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT WID_GRP_CD ""宽度组"", MINI ""宽度下限"", MAXI ""宽度上限"" FROM  NISCO.EP_ROLL_WID_GRP "
        DD.sQuery = DD.sQuery + " WHERE PLT || PRC_LINE = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND WID_GRP_CD   like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  WID_GRP_CD  ASC, MINI ASC,  MAXI ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT WID_GRP_CD ""宽度组"", MINI ""宽度下限"", MAXI ""宽度上限"" FROM  NISCO.EP_ROLL_WID_GRP "
        DD.sQuery = DD.sQuery + " WHERE PLT || PRC_LINE = '" & Trim(DD.sKey) & "' "
        DD.sWhere = DD.sWhere + "   AND WID_GRP_CD   like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  WID_GRP_CD  ASC, MINI ASC,  MAXI ASC "
        
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
'   1.ID           : Gf_PgmID_DD
'   2.Name         : Program ID Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2004. 2 .5
'   7.Modify Date  :
'   8.Comment      : Program ID Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_PgmID_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

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
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
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
    
    DD.DataDicType = "PGM"      'Program ID
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT PGMID ""程序 ID"", PGMNAME ""程序名称"" FROM  NISCO.ZP_PGMID "
        DD.sWhere = "             WHERE PGMID   like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  PGMID  ASC  "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT PGMID ""程序 ID"", PGMNAME ""程序名称"" FROM  NISCO.ZP_PGMID "
        DD.sWhere = "             WHERE PGMID   like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  PGMID  ASC  "
        
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
'   1.ID           : Gf_EmpID_DD
'   2.Name         : Employeed ID Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2004. 2 .5
'   7.Modify Date  :
'   8.Comment      : Employeed ID Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_EmpID_DD(Conn As ADODB.Connection, KeyCode As Integer, Optional pltType As String = "") As Boolean

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
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
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
    
    DD.DataDicType = "EMP"      'Program ID
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT EMP_ID ""人员 ID"", EMP_NAME ""人员名称"" FROM  NISCO.ZP_EMPLOYEE "
        DD.sWhere = "             WHERE EMP_ID   like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND DEPT     like '" & Trim(pltType) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  EMP_ID  ASC  "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
        
        DD.sQuery = "            SELECT EMP_ID ""人员 ID"", EMP_NAME ""人员名称"" FROM  NISCO.ZP_EMPLOYEE "
        DD.sWhere = "             WHERE EMP_ID   like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND DEPT     like '" & Trim(pltType) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  EMP_ID  ASC  "
        
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
'   1.ID           : Gf_heat_cond_DD
'   2.Name         : Order Usage Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Order Usage Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_HEAT_COND_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
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

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
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
    
    DD.DataDicType = "HC"          'HEAT　COND
    DD.DicRefType = "C"            'Active Form DataDic Call
     
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT HTM_COND ""热处理条件"", HTM_COND_TXT ""热处理条件说明"" ,HTM_TEMP_TGT  ""均热段温度"",  "
        DD.sQuery = DD.sQuery + "       HTM_TIME_1F_AIM ""均热段驻留时间(1)"",HTM_TIME_2F_AIM ""均热段驻留时间(2)"","
        DD.sQuery = DD.sQuery + "       HTM_COOL_TYP  ""冷却方式"", HTM_COOL_TMP  ""下冷床温度"""
        DD.sQuery = DD.sQuery + "  FROM NISCO.QP_HEAT_COND "
        DD.sWhere = "             WHERE HTM_COND             LIKE      '" & Trim(DD.sKey) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  HTM_COND  ASC  "
       
    Else
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
    
        DD.sQuery = "            SELECT HTM_COND ""热处理条件"", HTM_COND_TXT ""热处理条件说明"", HTM_TEMP_TGT  ""均热段温度"",  "
        DD.sQuery = DD.sQuery + "       HTM_TIME_1F_AIM ""均热段驻留时间(1)"",HTM_TIME_2F_AIM ""均热段驻留时间(2)"","
        DD.sQuery = DD.sQuery + "       HTM_COOL_TYP  ""冷却方式"", HTM_COOL_TMP  ""下冷床温度"""
        DD.sQuery = DD.sQuery + "  FROM NISCO.QP_HEAT_COND "
        DD.sWhere = "             WHERE HTM_COND             LIKE      '" & Trim(DD.sKey) & "%' "
        DD.sWhere = DD.sWhere + " ORDER  BY  HTM_COND  ASC  "
        
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
