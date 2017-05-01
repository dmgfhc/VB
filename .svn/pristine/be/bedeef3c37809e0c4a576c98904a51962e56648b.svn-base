Attribute VB_Name = "MasterCommon"
Option Explicit


'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_ExecQuery
'   2.Name         : Master Query Execute
'   3.Input  Value : vParam() Variant, Conn Connection, sQuery String
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Query Execute
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_ExecQuery(ByRef vParam() As Variant, Conn As ADODB.Connection, sQuery As String) As Boolean

On Error GoTo ExecQuery_ERROR

    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim adoCmd As ADODB.Command
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Ms_ExecQuery = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = Conn
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(vParam(1, 1), vParam(1, 2), vParam(1, 3), vParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(vParam(2, 1), vParam(2, 2), vParam(2, 3), vParam(2, 4))
    
    Conn.BeginTrans
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "0" Then
    
        Conn.RollbackTrans
        ret_Result_ErrCode = adoCmd("arg_e_code")
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
        
        Set adoCmd = Nothing
        Gf_Ms_ExecQuery = False
    
        Exit Function
        
    End If
    
    Conn.CommitTrans
    Set adoCmd = Nothing
    Gf_Ms_ExecQuery = True
    
    Exit Function

ExecQuery_ERROR:

    Conn.RollbackTrans
    Set adoCmd = Nothing
    Gf_Ms_ExecQuery = False
    
    sErrMessg = Err.Description & sQuery
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_MakeQuery
'   2.Name         : Master Make Query
'   3.Input  Value : ProcedureName Variant, iType String, {Retcol Collection}
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_MakeQuery(ProcedureName As Variant, iType As String, Optional Retcol As Collection) As String

On Error GoTo MasterMakeQuery_Error
   
    Dim iTemp_Int As Long
    Dim dTemp_Flo As Double
    Dim sQuery As String
    Dim sTemp As String
    Dim Ctrl As Control

    'Refer Or OneRow is No iType
    If iType = "R" Or iType = "O" Then
            sQuery = "{call " + ProcedureName + " ( "
        Else
            sQuery = "{call " + ProcedureName + " ( '" + iType + "',"
    End If

    If Not Retcol Is Nothing Then
    
        For Each Ctrl In Retcol
        
            If TypeOf Ctrl Is CheckBox Then
                If Ctrl = 1 Then
                    sQuery = sQuery + "'1',"
                Else
                    sQuery = sQuery + "'0',"
                End If
                
            ElseIf TypeOf Ctrl Is OptionButton Then
                If Ctrl = True Then
                    sQuery = sQuery + "'1',"
                Else
                    sQuery = sQuery + "'0',"
                End If
                
            ElseIf TypeOf Ctrl Is SSCheck Then
                If Ctrl = True Then
                    sQuery = sQuery + "'1',"
                Else
                    sQuery = sQuery + "'0',"
                End If
                
            ElseIf TypeOf Ctrl Is SSOption Then
                If Ctrl = True Then
                    sQuery = sQuery + "'1',"
                Else
                    sQuery = sQuery + "'0',"
                End If
                
            ElseIf TypeOf Ctrl Is ComboBox Then    'Modified by GuoLi
                If Ctrl.Style = 2 Then
                    If Ctrl.ListIndex = -1 Then
                        sQuery = sQuery + "'" & "" & "',"
                    Else
                        sQuery = sQuery + "'" + Trim(Ctrl.Text) + "',"
                    End If
                Else
                    sQuery = sQuery + "'" + Trim(Ctrl.Text) + "',"
                End If
                
            ElseIf TypeOf Ctrl Is silgEdit Then
                If Trim(Ctrl) = "" Then
                    iTemp_Int = 0
                Else
                    iTemp_Int = Ctrl.Value
                End If
                sQuery = sQuery & iTemp_Int & ","
                
            ElseIf TypeOf Ctrl Is sidbEdit Then
                If Trim(Ctrl) = "" Then
                    dTemp_Flo = 0
                Else
                    dTemp_Flo = Ctrl.Value
                End If
                sQuery = sQuery & dTemp_Flo & ","
                
            ElseIf TypeOf Ctrl Is sitxEdit Then
                If Ctrl = "____-__-__" Or Ctrl = "____-__" Or Ctrl = "____" Then
                    sQuery = sQuery + " '',"
                Else
                    sQuery = sQuery + "'" + Ctrl.RawData + "',"
                End If
            ElseIf TypeOf Ctrl Is sidtEdit Then
                If Ctrl = "____-__-__" Then
                    sQuery = sQuery + "'',"
                Else
                    sQuery = sQuery + "'" + Ctrl + "',"
                End If
            
            ElseIf TypeOf Ctrl Is UDate Then
                
                If Ctrl.MaxLength = 4 Then
                    sQuery = sQuery + "'" + Left(Trim(Ctrl.RawData), 4) + "',"
                ElseIf Ctrl.MaxLength = 7 Then
                    sQuery = sQuery + "'" + Mid(Trim(Ctrl.RawData), 1, 6) + "',"
                Else
                    sQuery = sQuery + "'" + Trim(Ctrl.RawData) + "',"
                End If
                
            ElseIf TypeOf Ctrl Is sitmEdit Then
                sTemp = Replace(Ctrl, "'", "''")
                sQuery = sQuery + "'" + sTemp + "',"
                
            ElseIf TypeOf Ctrl Is SSPanel Then
                sQuery = sQuery + "'" + Trim(Ctrl.Caption) + "',"
                
            ElseIf TypeOf Ctrl Is ULabel Then
                sQuery = sQuery + "'" + Trim(Ctrl.Caption) + "',"
                
            Else
                sTemp = Replace(Ctrl, "'", "''")
                sQuery = sQuery + "'" + Trim(sTemp) + "',"
            End If
                
        Next Ctrl
    End If
    
    'Refer Or OneRow is Last String Delete
    If iType = "R" Or iType = "O" Then
        sQuery = Mid(sQuery, 1, Len(sQuery) - 1) + ")}"
    Else
        sQuery = sQuery + "?,?)}"
    End If

    Gf_Ms_MakeQuery = sQuery
    
    Exit Function
    
MasterMakeQuery_Error:

    Gf_Ms_MakeQuery = "FAIL"
    sErrMessg = sQuery
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Ms_ControlLock
'   2.Name         : Control Lock
'   3.Input  Value : lControl Collection, Tf Boolean
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Control Lock
'---------------------------------------------------------------------------------------
Public Sub Gp_Ms_ControlLock(lControl As Collection, Tf As Boolean)
    
    Dim iCount As Integer

    For iCount = 1 To lControl.Count
        lControl.Item(iCount).Enabled = Not Tf
    Next iCount
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Ms_NeceColor
'   2.Name         : Control Necessary Color Setting
'   3.Input  Value : lControl Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Control Necessary Color Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Ms_NeceColor(lControl As Collection)
    
    Dim iCount As Integer

    For iCount = 1 To lControl.Count
        
        If TypeOf lControl.Item(iCount) Is DTPicker Then
            'lControl.Item(iCount).BackColor = &HC0FFFF
        Else
            lControl.Item(iCount).BackColor = &HC0FFFF
        End If
        
    Next iCount
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Ms_Collection
'   2.Name         : Master Collection Setting
'   3.Input  Value : Name Variant, pctl String, nctl String, mctl String, ictl String,
'                    rctl String, actl String, lctl String, Control Collection,
'                    nControl Collection, mControl Collection, iControl Collection,
'                    rControl Collection, aControl Collection, lControl Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Collection Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Ms_Collection(Name As Variant, pctl As String, nctl As String, mctl As String, ictl As String, _
                             rctl As String, actl As String, lctl As String, pControl As Collection, nControl As Collection, _
                             mControl As Collection, iControl As Collection, rControl As Collection, aControl As Collection, _
                             lControl As Collection)
    
    If LCase(Trim(pctl)) = "p" Then     'Primary Key Control
        pControl.Add Item:=Name
    End If
    
    If LCase(Trim(nctl)) = "n" Then     'Necessary Control
        nControl.Add Item:=Name
    End If
    
    If LCase(Trim(mctl)) = "m" Then     'Maxlength check Control
        mControl.Add Item:=Name
    End If
    
    If LCase(Trim(ictl)) = "i" Then     'Insert Control
        iControl.Add Item:=Name
    End If
    
    If LCase(Trim(rctl)) = "r" Then     'Refer Control
        rControl.Add Item:=Name
    End If
    
    If LCase(Trim(actl)) = "a" Then     'Master -> Spread Control
        aControl.Add Item:=Name
    End If
    
    If LCase(Trim(lctl)) = "l" Then     'Lock Control
        lControl.Add Item:=Name
    End If

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Ms_Cls
'   2.Name         : Master Control Clear Setting
'   3.Input  Value : Mc Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Clear Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Ms_Cls(MC As Collection)

    Dim Ct As Control
    Dim sCurDate As String
    
    sCurDate = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD') FROM DUAL")
    
    For Each Ct In MC
        If TypeOf Ct Is CheckBox Then               'CHECK BOX
            Ct.Value = False
        ElseIf TypeOf Ct Is OptionButton Then       'OPTION
            Ct.Value = False
        ElseIf TypeOf Ct Is SSCheck Then            '3D CHECK BOX
            Ct = False
        ElseIf TypeOf Ct Is SSOption Then           '3D OPTION
            Ct = False
        ElseIf TypeOf Ct Is ComboBox Then           'COMBO BOX
            If Ct.Style = 2 Then Ct.ListIndex = 0 Else Ct.Text = ""
        ElseIf TypeOf Ct Is sidbEdit Then           'CRECENT Float
            Ct.Text = 0
        ElseIf TypeOf Ct Is silgEdit Then           'CRECENT Integer
            Ct.Text = 0
        ElseIf TypeOf Ct Is sidtEdit Then           'CRECENT Date
            'Ct.Text = Format$(Now, "YYYY-MM-DD")
            Ct.Text = sCurDate
        ElseIf TypeOf Ct Is sitmEdit Then           'CRECENT Time
            Ct.Text = Format$(Now, "H:M:S")
        ElseIf TypeOf Ct Is sitxEdit Then           'CRECENT Test
            If Ct.Mask = "####/##/##" Then
                'Ct.Text = Format$(Now, "YYYY/MM/DD")
                Ct.Text = Mid(sCurDate, 1, 4) & "/" & Mid(sCurDate, 6, 2) & "/" & Mid(sCurDate, 9, 10)
            ElseIf Ct.Mask = "####/##" Then
                'Ct.Text = Format$(Now, "YYYY/MM")
                Ct.Text = Mid(sCurDate, 1, 4) & "/" & Mid(sCurDate, 6, 2)
            ElseIf Ct.Mask = "####" Then
                'Ct.Text = Format$(Now, "YYYY")
                Ct.Text = Mid(sCurDate, 1, 4)
            Else
                Ct.RawData = ""
            End If
        ElseIf TypeOf Ct Is UDate Then              'Indate
            '07.10.25 UPDATE BY KIM SUNG HO
            'Ct.RawData = Format(Now, "YYYYMMDD")
            '-----------------------------------------
            If Ct.MaxLength = 4 Then
                'Ct.RawData = Format(Now, "YYYY")
                Ct.RawData = Mid(sCurDate, 1, 4)
            ElseIf Ct.MaxLength = 7 Then
                'Ct.RawData = Format(Now, "YYYYMM")
                Ct.RawData = Mid(sCurDate, 1, 4) & Mid(sCurDate, 6, 2)
            Else
                'Ct.RawData = Format(Now, "YYYYMMDD")
                Ct.RawData = Mid(sCurDate, 1, 4) & Mid(sCurDate, 6, 2) & Mid(sCurDate, 9, 10)
            End If
            '-----------------------------------------
        ElseIf TypeOf Ct Is DTPicker Then           'VB Dtpicker
            'Ct.Value = Format(Now, "YYYY-MM-DD")
            Ct.Value = sCurDate
        ElseIf TypeOf Ct Is TextBox Then            'VB TEXT
            Ct.Text = ""
        ElseIf TypeOf Ct Is Picture Then            'PICTURE
            Ct.Picture = Nothing
        ElseIf TypeOf Ct Is Image Then              'IMAGE
            Ct.Picture = Nothing
        ElseIf TypeOf Ct Is ListBox Then            'List Box
            Ct.Clear
        ElseIf TypeOf Ct Is SSPanel Then            'SSPANEL
            Ct.Caption = ""
        ElseIf TypeOf Ct Is ULabel Then             'ULABEL
            Ct.Caption = ""
        Else
            Ct.Text = ""
        End If
        
    Next Ct

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Copy
'   2.Name         : Master Control Copy
'   3.Input  Value : Mc Collection
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Copy
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Copy(MC As Collection) As Boolean

On Error GoTo MasterCopy_Error

    Dim iCount As Integer

    'cControl Clear
    For iCount = 1 To MC.Item("cControl").Count
        MC.Item("cControl").Remove 1
    Next iCount

    'rControl --> cControl Copy
    For iCount = 1 To MC.Item("rControl").Count
    
        If TypeOf MC.Item("rControl").Item(iCount) Is CheckBox Then
            MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Value
        
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is OptionButton Then
            MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Value

        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSCheck Then
            MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Value
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSOption Then
            MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Value
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is ComboBox Then
            If MC.Item("rControl").Item(iCount).Style = 2 Then
                MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).ListIndex
            Else
                MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Text
            End If
        
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSPanel Then
            MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Caption
        
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is ULabel Then
            MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Caption
        
        Else
            MC.Item("cControl").Add Item:=MC.Item("rControl").Item(iCount).Text
        End If
        
    Next iCount

    Gf_Ms_Copy = True
    
    Exit Function
    
MasterCopy_Error:
    Gf_Ms_Copy = False

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_FormPaste
'   2.Name         : Form Control, Sheet Paste
'   3.Input  Value : Mc Collection, {Sc Collection}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Form Control, Sheet Paste
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_FormPaste(MC As Collection, Optional Sc As Collection) As Boolean

On Error GoTo FormPaste_Error

    Dim iCount As Integer

    'pControl(1) is Enabled=Ture Exit Function
    If MC!pControl(1).Enabled = True Then Gf_Ms_FormPaste = False: Exit Function
    
    'cControl.Count > 0 is Paste Process (cControl --> rControl)
    If MC("cControl").Count > 0 Then
        For iCount = 1 To MC.Item("cControl").Count
            If TypeOf MC.Item("rControl").Item(iCount) Is ComboBox Then             'COMBO BOX
                If MC.Item("rControl").Item(iCount).Style = 2 Then
                    MC.Item("rControl").Item(iCount).ListIndex = MC.Item("cControl").Item(iCount)
                Else
                    MC.Item("rControl").Item(iCount) = MC.Item("cControl").Item(iCount)
                End If
            ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSPanel Then         'COMBO BOX
                MC.Item("rControl").Item(iCount).Caption = MC.Item("cControl").Item(iCount)
            
            ElseIf TypeOf MC.Item("rControl").Item(iCount) Is ULabel Then          'COMBO BOX
                MC.Item("rControl").Item(iCount).Caption = MC.Item("cControl").Item(iCount)
            Else
                MC.Item("rControl").Item(iCount) = MC.Item("cControl").Item(iCount)
            End If
        Next iCount

        'Spread Check is True.....Spread copy --> Paste
        If Not Sc Is Nothing Then
        
            Call Gp_Sp_ClipCopy(Sc("Spread"), -1)
            Call Gp_Sp_ClipPaste(Sc("Spread"), -1)
            
            For iCount = 1 To Sc("Spread").MaxRows
                Sc("Spread").Col = 0: Sc("Spread").ROW = iCount
                Sc("Spread").Text = "Input"
            Next iCount
            
            Call Gp_Sp_CollectionLock(Sc("Spread"), Sc("pColumn"), False)
            
        End If
        
        Call Gp_Ms_ControlLock(MC("pcontrol"), False)
        Gf_Ms_FormPaste = True
        
    Else
        Gf_Ms_FormPaste = False
    End If
    
    Exit Function
    
FormPaste_Error:

    Gf_Ms_FormPaste = False
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Paste
'   2.Name         : Master Control Paste
'   3.Input  Value : Conn Connection, Mc Collection, {Sc Collection}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Paste
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Paste(Conn As ADODB.Connection, MC As Collection, Optional Sc As Collection) As Boolean

On Error GoTo MasterPaste_Error

    Dim iCount As Integer

    'pControl(1) is Enabled=Ture Exit Function
    If MC!pControl(1).Enabled = True Then Gf_Ms_Paste = False: Exit Function

    'cControl.Count > 0 is Paste Process (cControl --> rControl)
    If MC("cControl").Count > 0 Then
        For iCount = 1 To MC.Item("cControl").Count
        
            If TypeOf MC.Item("rControl").Item(iCount) Is ComboBox Then
                If MC.Item("rControl").Item(iCount).Style = 2 Then
                    MC.Item("rControl").Item(iCount).ListIndex = MC.Item("cControl").Item(iCount)
                Else
                    MC.Item("rControl").Item(iCount) = MC.Item("cControl").Item(iCount)
                End If
            ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSPanel Then
                MC.Item("rControl").Item(iCount).Caption = MC.Item("cControl").Item(iCount)
            ElseIf TypeOf MC.Item("rControl").Item(iCount) Is ULabel Then
                MC.Item("rControl").Item(iCount).Caption = MC.Item("cControl").Item(iCount)
            Else
                MC.Item("rControl").Item(iCount) = MC.Item("cControl").Item(iCount)
            End If
        Next iCount
        
        Call Gp_Ms_ControlLock(MC("pcontrol"), False)
        
        Gf_Ms_Paste = True
    Else
        Gf_Ms_Paste = False
    End If
    
    Exit Function
    
MasterPaste_Error:
    Gf_Ms_Paste = False
   
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_NeceCheck
'   2.Name         : Master Control Necessary Check
'   3.Input  Value : Retcol Collection
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Necessary Check
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_NeceCheck(Retcol As Collection) As String

    Dim II, i As Integer
    Dim YM As Integer
    YM = Retcol.Count
'
'    If Retcol.Count < 1 Then
'       Gf_Ms_NeceCheck = "OK"
'       Exit Function
'    End If
    
    For II = 1 To Retcol.Count
        If TypeOf Retcol.Item(II) Is CheckBox Then              'CHECK BOX
            If Retcol.Item(II).Value = False Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is OptionButton Then      'OPTION
            If Retcol.Item(II).Value = False Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If

        ElseIf TypeOf Retcol.Item(II) Is SSCheck Then           '3D CHECK BOX
            If Retcol.Item(II) = False Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is SSOption Then          '3D OPTION
            If Retcol.Item(II) = False Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is ComboBox Then          'COMBO BOX
            If Retcol.Item(II).Style = 2 Then
                If Retcol.Item(II).ListIndex = 0 Or Retcol.Item(II).ListIndex = -1 Then
                    Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                    Exit Function
                End If
            Else
                If Retcol.Item(II).Text = "" Then
                    Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                    Exit Function
                End If
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is sidbEdit Then          'sidbEdit
            If Retcol.Item(II).Value = Null Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is silgEdit Then          'silgEdit
            If Retcol.Item(II).Value = 0 Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is sitxEdit Or TypeOf Retcol.Item(II) Is sidtEdit Or TypeOf Retcol.Item(II) Is UDate Then
            For i = 1 To Retcol.Item(II).MaxLength
                If Mid(Retcol.Item(II).Text, i, 1) = Retcol.Item(II).FillChar Then
                    Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                    Exit Function
                End If
            Next i
            
        ElseIf TypeOf Retcol.Item(II) Is SSPanel Then          'SSPANEL
            If Retcol.Item(II).Caption = "" Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is ULabel Then           'ULABEL
            If Retcol.Item(II).Caption = "" Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
        
        Else
            If Retcol.Item(II).Text = "" Then
                Gf_Ms_NeceCheck = Retcol.Item(II).Tag
                Exit Function
            End If
        End If
        
    Next II
    
    Gf_Ms_NeceCheck = "OK"
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_NeceCheck2
'   2.Name         : Master Control Necessary, MaxLength Check
'   3.Input  Value : Retcol Collection
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Necessary, MaxLength Check
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_NeceCheck2(Retcol As Collection) As String

    Dim II, i As Integer
    
    For II = 1 To Retcol.Count
        If TypeOf Retcol.Item(II) Is CheckBox Then              'CHECK BOX
        
        ElseIf TypeOf Retcol.Item(II) Is OptionButton Then      'OPTION
            If Retcol.Item(II).Value = False Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If

        ElseIf TypeOf Retcol.Item(II) Is SSCheck Then           '3D CHECK BOX
            If Retcol.Item(II) = False Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is SSOption Then          '3D OPTION
            If Retcol.Item(II) = False Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is ComboBox Then          'COMBO BOX
            If Retcol.Item(II).Style = 2 Then
                If Retcol.Item(II).ListIndex = 0 Then
                    Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                    Exit Function
                End If
            Else
                If Retcol.Item(II).Text = "" Then
                    Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                    Exit Function
                End If
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is sidbEdit Then           'sidbEdit
            If Retcol.Item(II).Value = 0 Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is silgEdit Then            'silgEdit
            If Retcol.Item(II).Value = 0 Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is ListBox Then             'Listbox
            If Retcol.Item(II).Text = "" Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is sitxEdit Or TypeOf Retcol.Item(II) Is sidtEdit Or TypeOf Retcol.Item(II) Is UDate Then
            If Len(Trim(Retcol.Item(II).Text)) <> Retcol.Item(II).MaxLength Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            For i = 1 To Retcol.Item(II).MaxLength
                If Mid(Retcol.Item(II).Text, i, 1) = Retcol.Item(II).FillChar Then
                    Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                    Exit Function
                End If
            Next i
                
        ElseIf TypeOf Retcol.Item(II) Is SSPanel Then             'SSPanel
            If Retcol.Item(II).Caption = "" Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            
        ElseIf TypeOf Retcol.Item(II) Is ULabel Then              'ULabel
            If Retcol.Item(II).Caption = "" Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag
                Exit Function
            End If
            
        Else
            If Len(Trim(Retcol.Item(II).Text)) <> Retcol.Item(II).MaxLength Then
                Gf_Ms_NeceCheck2 = Retcol.Item(II).Tag + " = " + Trim(Str(Retcol.Item(II).MaxLength)) + " "
                Exit Function
            End If
        End If
    Next II
    
    Gf_Ms_NeceCheck2 = "OK"
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Refer
'   2.Name         : Master Control Refer
'   3.Input  Value : Conn Connection, Mc Collection, {nCheckControl Collection},
'                        {mCheckControl Collection}, {MsgChk Collection}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Refer
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Refer(Conn As ADODB.Connection, MC As Collection, Optional nCheckControl As Collection, _
                            Optional mCheckControl As Collection, Optional MsgChk As Boolean = True) As Boolean
    
On Error GoTo MasterRef_Err

    Dim sQuery As String
    Dim sMsg As String
    
    If Not MC Is Nothing Then
        If Not nCheckControl Is Nothing Then
            sMsg = Gf_Ms_NeceCheck(nCheckControl)
            If sMsg <> "OK" Then
                sMsg = sMsg + "必须输入"
                Call Gp_MsgBoxDisplay(sMsg)
                Gf_Ms_Refer = False
                Exit Function
            End If
        End If
        
        If Not mCheckControl Is Nothing Then
            sMsg = Gf_Ms_NeceCheck2(mCheckControl)
            If sMsg <> "OK" Then
                sMsg = sMsg + "长度不正确"
                Call Gp_MsgBoxDisplay(sMsg)
                Gf_Ms_Refer = False
                Exit Function
            End If
        End If
        
    End If
    
    'Make Query
    sQuery = Gf_Ms_MakeQuery(MC.Item("P-R"), "R", MC.Item("pControl"))
    
    If sQuery = "FAIL" Then
        Call Gp_MsgBoxDisplay("Refer Query Error : " & sErrMessg)
        Gf_Ms_Refer = False
        Exit Function
    End If
        
    'Query Excete and Display
    If Gf_Ms_Display(Conn, sQuery, MC.Item("rControl"), MC.Item("lControl")) = "OK" Then
        MDIMain.StatusBar1.Panels(1) = "提示信息：查询成功"
        Gf_Ms_Refer = True
    Else
        Gf_Ms_Refer = False
        If MsgChk Then
            Call Gp_MsgBoxDisplay("无相关记录", "I")
        End If
    End If
    
    Exit Function

MasterRef_Err:

    Call Gp_MsgBoxDisplay("Failed on data inquiry")
    Gf_Ms_Refer = False
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Display
'   2.Name         : Master Control Display
'   3.Input  Value : Conn Connection, sQuery String, Retcol Collection, Lockcon Collection
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Display
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Display(Conn As ADODB.Connection, sQuery As String, Retcol As Collection, Lockcon As Collection) As String

On Error GoTo MasterDisplay_Error
    
    Dim iCount As Integer
    Dim Atext As Variant
    Dim AdoRs As ADODB.Recordset
    
    Set AdoRs = New ADODB.Recordset
    
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Ms_Display = "FAIL": Exit Function
    End If

    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        While Not AdoRs.EOF
        
            If AdoRs(0) = "NOTHING" Then
                Gf_Ms_Display = "NOTHING"
                AdoRs.Close
                Set AdoRs = Nothing
                Exit Function
                
            ElseIf Mid(AdoRs(0), 1, 6) = "[FAIL]" Then
                Gf_Ms_Display = AdoRs(0)
                AdoRs.Close
                Set AdoRs = Nothing
                Exit Function
                
            End If
            
            For iCount = 1 To Retcol.Count
            
                If TypeOf Retcol.Item(iCount) Is CheckBox Then               'CHECK BOX
                    Retcol.Item(iCount).Value = IIf(VarType((AdoRs.Fields(iCount - 1))) = vbNull, "0", AdoRs.Fields(iCount - 1))
                    
                ElseIf TypeOf Retcol.Item(iCount) Is OptionButton Then       'OPTION
                    Retcol.Item(iCount).Value = IIf(VarType((AdoRs.Fields(iCount - 1))) = vbNull, 0, AdoRs.Fields(iCount - 1))
                    
                ElseIf TypeOf Retcol.Item(iCount) Is SSCheck Then            '3D CHECK BOX
                    Retcol.Item(iCount) = AdoRs.Fields(iCount - 1)
                    
                ElseIf TypeOf Retcol.Item(iCount) Is SSOption Then           '3D OPTION
                    Retcol.Item(iCount) = AdoRs.Fields(iCount - 1)
                    
                ElseIf TypeOf Retcol.Item(iCount) Is ComboBox Then           'COMBO BOX
                    If Retcol.Item(iCount).Style = 2 Then
                        If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                            Retcol.Item(iCount).ListIndex = 0
                        Else
                            Retcol.Item(iCount).ListIndex = Val(AdoRs.Fields(iCount - 1))
                        End If
                    Else
                        If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                            Retcol.Item(iCount).Text = ""
                        Else
                            Retcol.Item(iCount).Text = AdoRs.Fields(iCount - 1)
                        End If
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is sidbEdit Then           'sidbEdit
                    If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                        Retcol.Item(iCount).Value = 0
                    Else
                        Retcol.Item(iCount).Value = AdoRs.Fields(iCount - 1)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is silgEdit Then           'silgEdit
                    Retcol.Item(iCount).Value = AdoRs.Fields(iCount - 1)
                    If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                        Retcol.Item(iCount).Value = 0
                    Else
                        Retcol.Item(iCount).Value = AdoRs.Fields(iCount - 1)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is sidtEdit Then            'sidtEdit
                    If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                        Retcol.Item(iCount).Text = ""
                    Else
                        Retcol.Item(iCount).Text = AdoRs.Fields(iCount - 1)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is sitmEdit Then             'sitmEdit
                    If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                        Retcol.Item(iCount).Text = ""
                    Else
                        Retcol.Item(iCount).Text = AdoRs.Fields(iCount - 1)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is sitxEdit Then             'sitxEdit
                    If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                        Retcol.Item(iCount).Text = ""
                    Else
                        Retcol.Item(iCount).RawData = AdoRs.Fields(iCount - 1)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is DTPicker Then             'DTPicker
                    Atext = AdoRs.Fields(iCount - 1)
                    If VarType(Atext) = vbNull Then
                        Retcol.Item(iCount).Text = ""
                    Else
                        Retcol.Item(iCount).Value = Trim(Atext)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is UDate Then                'UDate
                    If VarType(AdoRs.Fields(iCount - 1)) = vbNull Then
                        Retcol.Item(iCount).RawData = ""
                    Else
                        Retcol.Item(iCount).RawData = AdoRs.Fields(iCount - 1)
                    End If
                
                ElseIf TypeOf Retcol.Item(iCount) Is TextBox Then              'TextBox
                    Atext = AdoRs.Fields(iCount - 1)
                    If VarType(Atext) = vbNull Then
                        Retcol.Item(iCount).Text = ""
                    Else
                        Retcol.Item(iCount).Text = Trim(Atext)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is SSPanel Then              'SSPanel
                    Atext = AdoRs.Fields(iCount - 1)
                    If VarType(Atext) = vbNull Then
                        Retcol.Item(iCount).Caption = ""
                    Else
                        Retcol.Item(iCount).Caption = Trim(Atext)
                    End If
                    
                ElseIf TypeOf Retcol.Item(iCount) Is ULabel Then               'ULabel
                    Atext = AdoRs.Fields(iCount - 1)
                    If VarType(Atext) = vbNull Then
                        Retcol.Item(iCount).Caption = ""
                    Else
                        Retcol.Item(iCount).Caption = Trim(Atext)
                    End If
                    
                Else
                    Atext = AdoRs.Fields(iCount - 1)
                    If VarType(Atext) = vbNull Then
                        Retcol.Item(iCount).Text = ""
                    Else
                        Retcol.Item(iCount).Text = Trim(Atext)
                    End If
                    
                End If
                
            Next iCount
            
            AdoRs.MoveNext
            
        Wend
    Else
        Gf_Ms_Display = ""
        Exit Function
    End If
    
    Call Gp_Ms_ControlLock(Lockcon, True)
    
    Set AdoRs = Nothing
    
    Gf_Ms_Display = "OK"
    Exit Function
    
MasterDisplay_Error:

    'Err.Raise AdoRs., Err.Description
    
    Set AdoRs = Nothing
    Gf_Ms_Display = "FAIL"
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Del
'   2.Name         : Master Control Delete
'   3.Input  Value : Conn Connection, Mc Collection
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Delete
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Del(Conn As ADODB.Connection, MC As Collection) As Boolean

On Error GoTo MasterDel_Error
    
    Dim iCount As Integer
    Dim sQuery As String
    Dim sMessg As String
    Dim OutParam(2, 4) As Variant
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256

    Gf_Ms_Del = True
    
    'pControl Enabled=true is Not Delete
    For iCount = 1 To MC.Item("pControl").Count
        If MC.Item("pControl")(iCount).Enabled = True Then
            Call Gp_MsgBoxDisplay("Inquire First Data", "I")
            Gf_Ms_Del = False
            Exit Function
        End If
    Next iCount
    
    'delete Confirm Message
    If Not Gf_MessConfirm("您确定要删除当前数据吗？", "Q") Then Exit Function
    
    'Delete Make Query
    sQuery = Gf_Ms_MakeQuery(MC.Item("P-M"), "D", MC.Item("iControl"))
    
    If sQuery = "FAIL" Then
        Call Gp_MsgBoxDisplay("Delete Query Error : " & sErrMessg)
        Gf_Ms_Del = False
        Exit Function
    End If

    'sMessg = Gf_Ms_Display(Conn, sQuery, Mc.Item("rControl"), Mc.Item("lControl"))
    
    'Query Process
    If Gf_Ms_ExecQuery(OutParam, Conn, sQuery) Then
        Call Gp_Ms_ControlLock(MC!pControl, False)
        MDIMain.StatusBar1.Panels(1) = "提示信息：数据删除成功"
        Gf_Ms_Del = True
    Else
        Gf_Ms_Del = False
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Exit Function

MasterDel_Error:

    Gf_Ms_Del = False

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Process
'   2.Name         : Master Control Process
'   3.Input  Value : Conn Connection, Mc Collection, sAuthority String
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master Control Process
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Process(Conn As ADODB.Connection, MC As Collection, sAuthority As String) As Boolean

On Error GoTo MasterPro_Error

    Dim II As Integer
    Dim sQuery As String
    Dim sWhere As String
    Dim sMessg As String
    Dim OutParam(2, 4) As Variant
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256

    'Necessarily Check
    sMessg = Gf_Ms_NeceCheck(MC.Item("nControl"))
    
    If Trim(sMessg) <> "OK" Then
        Call Gp_MsgBoxDisplay(Trim(sMessg) + "必须输入", "I")
        Gf_Ms_Process = False
        Exit Function
    End If

    'Maxlength Check
    sMessg = Gf_Ms_NeceCheck2(MC.Item("mControl"))
    
    If Trim(sMessg) <> "OK" Then
        Call Gp_MsgBoxDisplay(Trim(sMessg) + "长度不正确", "I")
        Gf_Ms_Process = False
        Exit Function
    End If
    
    If MC!pControl.Count > 0 And MC.Item("pControl")(1).Enabled = True Then
    
        'Insert Make Query
        sQuery = Gf_Ms_MakeQuery(MC.Item("P-M"), "I", MC.Item("iControl"))
        
        If sQuery = "FAIL" Then
            Call Gp_MsgBoxDisplay("Insert Query Error : " & sErrMessg)
            Gf_Ms_Process = False
            Exit Function
        End If
        
        If Gf_Ms_ExecQuery(OutParam, Conn, sQuery) Then
        
            sQuery = Gf_Ms_MakeQuery(MC.Item("P-R"), "R", MC.Item("pControl"))
            
            If sQuery = "FAIL" Then
                Call Gp_MsgBoxDisplay("Refer Query Error : " & sErrMessg)
                Gf_Ms_Process = False
                Exit Function
            End If
            
            Call Gf_Ms_Display(Conn, sQuery, MC!rControl, MC!lControl)
            Call Gp_Ms_ControlLock(MC!pControl, True)
            Gf_Ms_Process = True
            MDIMain.StatusBar1.Panels(1) = "提示信息：新增数据成功"
        Else
            Gf_Ms_Process = False
            Call Gp_MsgBoxDisplay(sErrMessg)
        End If
        
    Else
    
        If Mid(sAuthority, 3, 1) = "0" Then Gf_Ms_Process = True: Exit Function
        
        'Update Make Query
        sQuery = Gf_Ms_MakeQuery(MC.Item("P-M"), "U", MC.Item("iControl"))
        
        If sQuery = "FAIL" Then
            Call Gp_MsgBoxDisplay("Modify Query Error : " & sErrMessg)
            Gf_Ms_Process = False
            Exit Function
        End If
        
        If Gf_Ms_ExecQuery(OutParam, Conn, sQuery) Then
        
            sQuery = Gf_Ms_MakeQuery(MC.Item("P-R"), "R", MC.Item("pControl"))
            
            If sQuery = "FAIL" Then
                Call Gp_MsgBoxDisplay("Refer Query Error : " & sErrMessg)
                Gf_Ms_Process = False
                Exit Function
            End If
        
            Call Gf_Ms_Display(Conn, sQuery, MC!rControl, MC!lControl)
            Gf_Ms_Process = True
            MDIMain.StatusBar1.Panels(1) = "提示信息：数据更新成功"
        Else
            Gf_Ms_Process = False
            Call Gp_MsgBoxDisplay(sErrMessg)
        End If
        
    End If
    
    Exit Function
    
MasterPro_Error:

    Gf_Ms_Process = False
    Call Gp_MsgBoxDisplay("Failed in data processing")
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Outpara
'   2.Name         : Master Control Out Parameter
'   3.Input  Value : Conn Connection, Mc Collection, {RefChek Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 07 .15
'   7.Modify Date  :
'   8.Comment      : Master Control Out Parameta
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Outpara(Conn As ADODB.Connection, MC As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo Outpara_Error

    Dim iCount As Integer
    
    Dim dTempInt As Double
    
    Dim sMesg As String
    Dim sTemp As String
    Dim sQuery As String
    Dim Atext As Variant
    
    Dim adoCmd As ADODB.Command

    Gf_Ms_Outpara = True
    
    Screen.MousePointer = vbHourglass
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Ms_Outpara = False: Exit Function
    End If
    
    sQuery = Gf_Ms_MakeQuery(MC.Item("P-R"), "R", MC.Item("pControl"))
    
    sQuery = Mid(sQuery, 1, Len(Trim(sQuery)) - 2)
    
    For iCount = 1 To MC.Item("rControl").Count
        sQuery = sQuery + ",?"
    Next iCount
    
    sQuery = sQuery + ")}"
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = Conn
    
    adoCmd.CommandText = sQuery
    
    'Ceate Parameter (Output)
    For iCount = 1 To MC.Item("rControl").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount - 1), adVariant, adParamOutput)
    Next iCount
    
    adoCmd.Execute , , adExecuteNoRecords
    
    For iCount = 1 To MC.Item("rControl").Count
        
        If TypeOf MC.Item("rControl").Item(iCount) Is CheckBox Then                  'CHECK BOX
            MC.Item("rControl").Item(iCount).Value = IIf(VarType((adoCmd(Str(iCount - 1)))) = vbNull, "0", adoCmd(Str(iCount - 1)))
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is OptionButton Then          'OPTION
            MC.Item("rControl").Item(iCount).Value = IIf(VarType(adoCmd(Str(iCount - 1))) = vbNull, 0, adoCmd(Str(iCount - 1)))
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSCheck Then               '3D CHECK BOX
            MC.Item("rControl").Item(iCount) = adoCmd(Str(iCount - 1))
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSOption Then              '3D OPTION
            MC.Item("rControl").Item(iCount) = adoCmd(Str(iCount - 1))
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is ComboBox Then              'COMBO BOX
            If MC.Item("rControl").Item(iCount).Style = 2 Then
                If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                    MC.Item("rControl").Item(iCount).ListIndex = 0
                Else
                    MC.Item("rControl").Item(iCount).ListIndex = Val(adoCmd(Str(iCount - 1)))
                End If
            Else
                If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                    MC.Item("rControl").Item(iCount).Text = ""
                Else
                    MC.Item("rControl").Item(iCount).Text = adoCmd(Str(iCount - 1))
                End If
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is sidbEdit Then             'sidbEdit
            If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                MC.Item("rControl").Item(iCount).Value = 0
            Else
                MC.Item("rControl").Item(iCount).Value = adoCmd(Str(iCount - 1))
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is silgEdit Then             'silgEdit
            MC.Item("rControl").Item(iCount).Value = adoCmd(Str(iCount - 1))
            If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                MC.Item("rControl").Item(iCount).Value = 0
            Else
                MC.Item("rControl").Item(iCount).Value = adoCmd(Str(iCount - 1))
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is sidtEdit Then             'sidtEdit
            If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                MC.Item("rControl").Item(iCount).Text = ""
            Else
                MC.Item("rControl").Item(iCount).Text = adoCmd(Str(iCount - 1))
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is sitmEdit Then             'sitmEdit
            If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                MC.Item("rControl").Item(iCount).Text = ""
            Else
                MC.Item("rControl").Item(iCount).Text = adoCmd(Str(iCount - 1))
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is sitxEdit Then             'sitxEdit
            If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                MC.Item("rControl").Item(iCount).Text = ""
            Else
                MC.Item("rControl").Item(iCount).RawData = adoCmd(Str(iCount - 1))
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is DTPicker Then             'DTPicker
            Atext = adoCmd(Str(iCount - 1))
            If VarType(Atext) = vbNull Then
                MC.Item("rControl").Item(iCount).Text = ""
            Else
                MC.Item("rControl").Item(iCount).Value = Trim(Atext)
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is UDate Then                'UDate
            If VarType(adoCmd(Str(iCount - 1))) = vbNull Then
                MC.Item("rControl").Item(iCount).RawData = ""
            Else
                MC.Item("rControl").Item(iCount).RawData = adoCmd(Str(iCount - 1))
            End If
        
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is TextBox Then              'TextBox
            Atext = adoCmd(Str(iCount - 1))
            If VarType(Atext) = vbNull Then
                MC.Item("rControl").Item(iCount).Text = ""
            Else
                MC.Item("rControl").Item(iCount).Text = Trim(Atext)
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is SSPanel Then              'SSPanel
            Atext = adoCmd(Str(iCount - 1))
            If VarType(Atext) = vbNull Then
                MC.Item("rControl").Item(iCount).Caption = ""
            Else
                MC.Item("rControl").Item(iCount).Caption = Trim(Atext)
            End If
            
        ElseIf TypeOf MC.Item("rControl").Item(iCount) Is ULabel Then               'ULabel
            Atext = adoCmd(Str(iCount - 1))
            If VarType(Atext) = vbNull Then
                MC.Item("rControl").Item(iCount).Caption = ""
            Else
                MC.Item("rControl").Item(iCount).Caption = Trim(Atext)
            End If
            
        Else
            Atext = adoCmd(Str(iCount - 1))
            If VarType(Atext) = vbNull Then
                MC.Item("rControl").Item(iCount).Text = ""
            Else
                MC.Item("rControl").Item(iCount).Text = Trim(Atext)
            End If
            
        End If
        
    Next iCount
    
    Screen.MousePointer = vbDefault
    
    Set adoCmd = Nothing
    Gf_Ms_Outpara = True
    
    Exit Function
    
Outpara_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Gf_Ms_Outpara = False
    
    Err.Raise Err.Number, Err.Description

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_AllDel
'   2.Name         : Master-Sheet All Delete Process
'   3.Input  Value : Conn Connection, Sc Collection, {Mc Collection}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Master-Sheet All Delete Process
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_AllDel(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection) As Boolean

On Error GoTo AllDel_Error
    
    Dim sQuery As String
    Dim sMesg As String
    Dim sErrorID As String
    Dim sDel_MSG As String
    
    Dim OutParam(2, 4) As Variant

    Dim iCount As Integer
    Dim iProcessCount As Integer
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256

    iProcessCount = 0
        
    If Not MC Is Nothing Then
        'PK Lock Check
        For iCount = 1 To MC("pControl").Count
        
            If MC("pControl").Item(iCount).Enabled Then
                Call Gp_MsgBoxDisplay("Inquire first data", "I")
                Gf_Ms_AllDel = False
                Exit Function
            End If
            
        Next iCount
        
    End If
    
    'Delete Check Confirm
    If Gf_MessConfirm("您确定要删除这些数据吗？", "Q") = False Then Gf_Ms_AllDel = True: Exit Function
    
    If Sc.Item("Spread").MaxRows < 1 Then
    Else
        If Gf_Sp_DelProcess(Conn, Sc) = False Then
            Gf_Ms_AllDel = False
            Exit Function
        End If
    End If
    
    'Header Delete Process
    sQuery = Gf_Ms_MakeQuery(MC.Item("P-M"), "D", MC.Item("iControl"))
    
    If sQuery = "FAIL" Then
        Call Gp_MsgBoxDisplay("Delete Query Error : " & sErrMessg)
        Gf_Ms_AllDel = False
        Exit Function
    End If
        
    If Gf_Ms_ExecQuery(OutParam, Conn, sQuery) Then
    
        Call Gp_Ms_ControlLock(MC!pControl, False)
        
        If Sc.Item("Spread").MaxRows < 1 Then
        Else
            For iCount = 1 To Sc.Item("Spread").MaxRows
                Sc.Item("Spread").ROW = iCount: Sc.Item("Spread").Col = 0
                Sc.Item("Spread").Text = "Input"
            Next iCount
            
            Call Gp_Sp_CollectionLock(Sc.Item("Spread"), Sc.Item("pColumn"), False)
        End If
        
        MDIMain.StatusBar1.Panels(1) = "提示信息：数据删除成功"
        Gf_Ms_AllDel = True
        
    Else
        Gf_Ms_AllDel = False
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
        
    Exit Function
    
AllDel_Error:

    Gf_Ms_AllDel = False
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Ms_Rset
'   2.Name         : RecordSet Value Rerurn
'   3.Input  Value : Conn Connection, sQuery String
'   4.Return Value : Recordset
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2004. 02 .05
'   7.Modify Date  :
'   8.Comment      : RecordSet Value Return
'   9.Use Method   : TreeView Control Use
'---------------------------------------------------------------------------------------
Public Function Gf_Ms_Rset(Conn As ADODB.Connection, sQuery As String) As ADODB.Recordset

On Error GoTo RecordSet_Error

    Dim AdoRs As ADODB.Recordset
    Set AdoRs = New ADODB.Recordset

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Exit Function
    End If
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        If Not AdoRs.EOF Then
            Set Gf_Ms_Rset = AdoRs
        End If
    Else
        AdoRs.Close
    End If
    
    Exit Function
    
RecordSet_Error:
    Set AdoRs = Nothing
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_ComnNameFind
'   2.Name         : Common Code Name Return
'   3.Input  Value : Conn Connection, Cd_Mana_No String, Code String, nameType String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Common Code Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_ComnNameFind(Conn As ADODB.Connection, Cd_Mana_No As String, Code As String, nameType As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_ComnNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    Select Case nameType
    
        Case "1"        'Short Name
            sQuery = "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = '" & Cd_Mana_No & "' AND CD = '" & Code & "' "
        Case "2"        'Full Name
            sQuery = "SELECT CD_NAME       FROM ZP_CD WHERE CD_MANA_NO = '" & Cd_Mana_No & "' AND CD = '" & Code & "' "
        Case "3"        'Short Eng Name
            sQuery = "SELECT CD_SHORT_ENG  FROM ZP_CD WHERE CD_MANA_NO = '" & Cd_Mana_No & "' AND CD = '" & Code & "' "
        Case "4"        'Full Eng Name
            sQuery = "SELECT CD_FULL_ENG   FROM ZP_CD WHERE CD_MANA_NO = '" & Cd_Mana_No & "' AND CD = '" & Code & "' "
        Case Else       'Full Name
            sQuery = "SELECT CD_NAME       FROM ZP_CD WHERE CD_MANA_NO = '" & Cd_Mana_No & "' AND CD = '" & Code & "' "
            
    End Select
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_ComnNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_ComnNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_ComnNameFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_UsageNameFind
'   2.Name         : Ord Usage Code Name Return
'   3.Input  Value : Conn Connection, Prod_Knd String, Code String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Ord Usage Code Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_UsageNameFind(Conn As ADODB.Connection, Prod_Knd As String, Code As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_UsageNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
        
    'Name
    sQuery = "SELECT ENDUSE_NAME FROM QP_ORD_USAGE WHERE PROD_KND = '" & Prod_Knd & "' AND ENDUSE_CD = '" & Code & "' "
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_UsageNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_UsageNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_UsageNameFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_CustNameFind
'   2.Name         : Customer Name Return
'   3.Input  Value : Conn Connection, Code String, nameType String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Customer Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_CustNameFind(Conn As ADODB.Connection, Code As String, nameType As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_CustNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    Select Case nameType
    
        Case "1"        'Name
            sQuery = "SELECT CUST_NM      FROM BP_CUST_CD WHERE CUST_CD = '" & Code & "' "
        Case "2"        'Eng Name
            sQuery = "SELECT CUST_NM_ENG  FROM BP_CUST_CD WHERE CUST_CD = '" & Code & "' "
        Case Else       'Name
            sQuery = "SELECT CUST_NM      FROM BP_CUST_CD WHERE CUST_CD = '" & Code & "' "
            
    End Select
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_CustNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_CustNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_CustNameFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_DestNameFind
'   2.Name         : Destination Name Return
'   3.Input  Value : Conn Connection, Code String, nameType String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Destination Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_DestNameFind(Conn As ADODB.Connection, Code As String, nameType As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_DestNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    Select Case nameType
    
        Case "1"        'Name
            sQuery = "SELECT DEST_NM      FROM BP_DEST_CD WHERE DEST_CD = '" & Code & "' "
        Case "2"        'Eng Name
            sQuery = "SELECT DEST_NM_ENG  FROM BP_DEST_CD WHERE DEST_CD = '" & Code & "' "
        Case Else
            sQuery = "SELECT DEST_NM      FROM BP_DEST_CD WHERE DEST_CD = '" & Code & "' "
        
    End Select
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_DestNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_DestNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_DestNameFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_CarInfFind
'   2.Name         : Common Code Name Return
'   3.Input  Value : Conn Connection, Car_no String, Car_knd String, nameType String
'   4.Return Value : Variant
'   5.Writer       : Li Chao
'   6.Create Date  : 2012. 07 .23
'   7.Modify Date  :
'   8.Comment      : Common Code Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_CarInfFind(Conn As ADODB.Connection, Car_no As String, Car_knd As String, nameType As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_CarInfFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    Select Case nameType
    
        Case "1"        '最大装载量
            sQuery = "SELECT H.CAR_WGT_MAX    FROM HP_CAR_IMF H WHERE H.CAR_NO = '" & Car_no & "' AND H.CAR_KND = '" & Car_knd & "' "
        Case "2"        '装载量(适量)
            sQuery = "SELECT H.CAR_WGT_AVE    FROM HP_CAR_IMF H WHERE H.CAR_NO = '" & Car_no & "' AND H.CAR_KND = '" & Car_knd & "' "
        Case Else       '可装车长度
            sQuery = "SELECT H.CAR_LEN        FROM HP_CAR_IMF H WHERE H.CAR_NO = '" & Car_no & "' AND H.CAR_KND = '" & Car_knd & "' "
            
    End Select
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_CarInfFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_CarInfFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_CarInfFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_ApplyNameFind
'   2.Name         : Apply Code Name Return
'   3.Input  Value : Conn Connection, Table_id String, Code String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .19
'   7.Modify Date  :
'   8.Comment      : Common Code Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_ApplyNameFind(Conn As ADODB.Connection, Table_id As String, Code As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_ApplyNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    sQuery = "SELECT APLY_ITEM_NAME FROM ZP_APLY_ITEM WHERE APLY_ITEM = '" & Code & "' AND TABLE_ID = '" & Table_id & "' "
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_ApplyNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_ApplyNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_ApplyNameFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_StlgrdNameFind
'   2.Name         : Apply Code Name Return
'   3.Input  Value : Conn Connection, Code String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 07 .24
'   7.Modify Date  :
'   8.Comment      : Stlgrd Code Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_StlgrdNameFind(Conn As ADODB.Connection, Code As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_StlgrdNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    sQuery = "SELECT STEEL_GRD_DETAIL  FROM QP_NISCO_CHMC WHERE STLGRD = '" & Code & "' "
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_StlgrdNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_StlgrdNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_StlgrdNameFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_PgmNameFind
'   2.Name         : Program ID Name Return
'   3.Input  Value : Conn Connection, Code String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2004. 02 .5
'   7.Modify Date  :
'   8.Comment      : Program ID Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_PgmNameFind(Conn As ADODB.Connection, Code As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_PgmNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    sQuery = "SELECT PGMNAME FROM ZP_PGMID WHERE PGMID = '" & Code & "' "
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_PgmNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_PgmNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_PgmNameFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_EmpNameFind
'   2.Name         : Employeed ID Name Return
'   3.Input  Value : Conn Connection, Code String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2004. 02 .5
'   7.Modify Date  :
'   8.Comment      : Employeed ID Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_EmpNameFind(Conn As ADODB.Connection, Code As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_EmpNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    sQuery = "SELECT EMP_NAME FROM ZP_EMPLOYEE WHERE EMP_ID = '" & Code & "' "
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_EmpNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_EmpNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_EmpNameFind = "FAIL"

End Function

