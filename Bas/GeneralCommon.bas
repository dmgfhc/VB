Attribute VB_Name = "GeneralCommon"
Option Explicit

Public M_CN1 As New ADODB.Connection     'Connection
Public AdoRs As ADODB.Recordset          'Record Set
Public adoCmd As ADODB.Command           'Command

Public Active_Spread As Object           'Spread Object

Public sUserName As String               'User Name
Public sUserID As String                 'User Id
Public PassCheck As Boolean              'Password Check
Public sErrMessg As String               'Error Message

Public iDupCnt As Integer                'Duplicate exclusion Count
Public iSumCnt As Integer                'Sum Column Count

Public MainFrmType As String              ' 2012.11.09 新增  耿朝雷

Type DataDic
    sKey As String                       'Condition Key
    sQuery As String                     'sQuery
    sWhere As String                     'sWhere
    sWitch As String                     'Control, Spread Type
    nameType As String                   'Name Type
    DicRefType As String                 'DataDic Refer Type
    DataDicType As String                'DataDic Type
    sSelect As Boolean                   'Data Select Status
    sPname As vaSpread                   'Spread Name
    rControl As New Collection           'Control (Code, Name)
    wControl As New Collection           'Where Control
End Type
 
Public DD As DataDic

'---------------------------------------------------------------------------------------
'   1.ID           : GF_DbConnect
'   2.Name         : DataBase Connection
'   3.Input  Value :
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : DataBase Connection
'---------------------------------------------------------------------------------------
Public Function GF_DbConnect() As Boolean
'
On Error GoTo DbConnect_ERROR

    Screen.MousePointer = vbHourglass
    
'    M_CN1.ConnectionString = "Provider=MSDAORA.1;User ID=nisco/nisco01;Data Source=web;Persist Security Info=True"

    M_CN1.ConnectionString = "Provider=MSDAORA.1;User ID=nisco/nisco01;Data Source=ora9;Persist Security Info=True"
    
    M_CN1.CursorLocation = adUseClient
    
    M_CN1.CommandTimeout = 10
        
    M_CN1.ConnectionTimeout = 10
     
    M_CN1.Open
    
    GF_DbConnect = True
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
DbConnect_ERROR:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("数据库连接失败，请稍后再试")
    GF_DbConnect = False
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_MessConfirm
'   2.Name         : Message Box Confirm
'   3.Input  Value : sMsg String, {sIcon String}, {sTitle String}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Message Box Confirm Value Return
'---------------------------------------------------------------------------------------
Public Function Gf_MessConfirm(ByVal sMsg As String, Optional sIcon As String, Optional sTitle As String) As Boolean
    
    Dim sStyle As String
    Dim iRet As Integer
    
    'Message Box Style Selection
    Select Case sIcon
        Case "Q"
            sStyle = vbYesNo + vbQuestion + vbDefaultButton2
        Case "W"
            sStyle = vbYesNo + vbExclamation + vbDefaultButton2
        Case "I"
            sStyle = vbYesNo + vbInformation + vbDefaultButton2
        Case Else
            sStyle = vbYesNo + vbCritical + vbDefaultButton2
    End Select
    
    If RTrim(sTitle) = "" Then
        sTitle = "系统提示信息确认"
    End If

    iRet = MsgBox(sMsg, sStyle, sTitle)
    
    If iRet = vbYes Then
        Gf_MessConfirm = True
    Else
        Gf_MessConfirm = False
    End If
        
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_MsgBoxDisplay
'   2.Name         : Message Box Display
'   3.Input  Value : sMsg String, {sIcon String}, {sTitle String}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Message Box Only Display
'---------------------------------------------------------------------------------------
Public Sub Gp_MsgBoxDisplay(ByVal sMsg As String, Optional sIcon As String, Optional sTitle As String)
    
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
        Case Else
            sStyle = vbOKOnly + vbCritical
    End Select
    
    If RTrim(sTitle) = "" Then
        sTitle = "系统提示信息"
    End If

    Call MsgBox(sMsg, sStyle, sTitle)
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_FloatFind
'   2.Name         : Float Value Return
'   3.Input  Value : Conn Connection, sQuery String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Float Value Return
'---------------------------------------------------------------------------------------
Public Function Gf_FloatFind(Conn As ADODB.Connection, sQuery As String) As Variant

On Error GoTo FloatFind_Error

    Dim AdoRs As ADODB.Recordset
    
    Set AdoRs = New ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_FloatFind = 0: Exit Function
    End If

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset

    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_FloatFind = 0
            Else
                Gf_FloatFind = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_FloatFind = 0
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

FloatFind_Error:

    Set AdoRs = Nothing
    Gf_FloatFind = 0

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_CodeFind
'   2.Name         : Code Name Return
'   3.Input  Value : Conn Connection, sQuery String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Text Code Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_CodeFind(Conn As ADODB.Connection, sQuery As String) As Variant

On Error GoTo CodeFind_Error

    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_CodeFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_CodeFind = ""
            Else
                Gf_CodeFind = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_CodeFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_CodeFind = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant, sQuery String, {ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 07 .14
'   7.Modify Date  :
'   8.Comment      : combo Add
'---------------------------------------------------------------------------------------
Public Function Gf_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sQuery As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_ComboAdd = False: Exit Function
    End If
    
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
        Gf_ComboAdd = True
    Else
        Gf_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_ComboAdd = False

End Function

Public Function Gf_ComboAdd2(Conn As ADODB.Connection, Cbo As Variant, sQuery As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo Gf_ComboAdd2_Error
    
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_ComboAdd2 = False: Exit Function
    End If
    
    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset
     
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
     
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
               If AdoRs.Fields(1) > "0" Then
                  
                  Cbo.AddItem AdoRs.Fields(0) + " E"
                  
               Else
                  Cbo.AddItem AdoRs.Fields(0)
              
              
               End If
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_ComboAdd2 = True
    Else
        Gf_ComboAdd2 = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

Gf_ComboAdd2_Error:

    Set AdoRs = Nothing
    Gf_ComboAdd2 = False

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_FormCenter
'   2.Name         : Form Center
'   3.Input  Value : Fm Variant
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Form Center
'---------------------------------------------------------------------------------------
Public Sub Gp_FormCenter(Fm As Variant)

    Fm.Left = (Screen.Width - Fm.Width) / 2
    Fm.Top = (Screen.Height - Fm.Height - 650) / 2
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_FormLoc_Get
'   2.Name         : Form Location Get
'   3.Input  Value : oForm Form, {fType tring}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Form Location Get
'---------------------------------------------------------------------------------------
Public Sub Gp_FormLoc_Get(oForm As Form, Optional fType As String = "")

    Dim sEcname As String
    Dim sFileName As String
    Dim sKey As String

    sEcname = oForm.Name + fType
    sFileName = "Z-SYSTEM.INI"
    
    sKey = "TOP"
    oForm.Top = GetPrivateProfileInt(sEcname, sKey, oForm.Top, App.Path & "\" & sFileName)
    
    sKey = "LEFT"
    oForm.Left = GetPrivateProfileInt(sEcname, sKey, oForm.Left, App.Path & "\" & sFileName)
    
    sKey = "HEIGHT"
    oForm.Height = GetPrivateProfileInt(sEcname, sKey, oForm.Height, App.Path & "\" & sFileName)

    sKey = "WIDTH"
    oForm.Width = GetPrivateProfileInt(sEcname, sKey, oForm.Width, App.Path & "\" & sFileName)

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_FormLoc_Set
'   2.Name         : Form Location Set
'   3.Input  Value : oForm Form, {fType String}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Form Location Set
'---------------------------------------------------------------------------------------
Public Sub Gp_FormLoc_Set(oForm As Form, Optional fType As String = "")

    Dim sEcname As String
    Dim sFileName As String
    Dim sKey As String
    Dim sValue As String
    
    sEcname = oForm.Name + fType
    sFileName = "Z-SYSTEM.INI"

    sKey = "TOP": sValue = oForm.Top
    Call WritePrivateProfileString(sEcname, sKey, sValue, App.Path & "\" & sFileName)
    
    sKey = "LEFT": sValue = oForm.Left
    Call WritePrivateProfileString(sEcname, sKey, sValue, App.Path & "\" & sFileName)
    
    sKey = "HEIGHT": sValue = oForm.Height
    Call WritePrivateProfileString(sEcname, sKey, sValue, App.Path & "\" & sFileName)
    
    sKey = "WIDTH": sValue = oForm.Width
    Call WritePrivateProfileString(sEcname, sKey, sValue, App.Path & "\" & sFileName)

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_DateSetting
'   2.Name         : Client PC Data Format Setting (Register)
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Client PC Data Format Setting (Register)
'---------------------------------------------------------------------------------------
Public Sub Gp_DateSetting()

    Call Gp_SetRegValue(&H80000001, "Control Panel\International", "sShortDate", "yyyy-MM-dd", 1)

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_SetRegValue
'   2.Name         : Date Register Open
'   3.Input  Value : lhKey Long, sKeyName String, sValueName String,
'                    vValueSetting Variant, lValueType Long
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Date Register Open
'---------------------------------------------------------------------------------------
Public Sub Gp_SetRegValue(ByVal lhKey As Long, sKeyName As String, sValueName As String, _
                          vValueSetting As Variant, lValueType As Long)
    
    Dim lRetVal As Long
    Dim hKey As Long

    'Open the Specified Key
    lRetVal = RegOpenKeyEx(lhKey, sKeyName, 0, &H2, hKey)
    lRetVal = Gf_SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_SetValueEx
'   2.Name         : Setting Register Date Format
'   3.Input  Value : hKey Long, sValueName String, lType Long, vValue Variant
'   4.Return Value : Long
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Setting Register Date Format
'---------------------------------------------------------------------------------------
Public Function Gf_SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    
    Dim I As Integer
    Dim j As Integer
    Dim strValue As String
    Dim lngValue As Long

    j = 0

    Select Case lType
    
        Case 1
            strValue = vValue
            
            For I = 1 To Len(strValue)
                If Asc(Mid(strValue, I, 1)) < 0 Then j = j + 1
            Next I
            
            If j = 0 Then
                I = Len(strValue)
            Else
                I = LenB(strValue) - (Len(strValue) - j)
            End If
            
            Gf_SetValueEx = RegSetValueExString(hKey, sValueName, 0, lType, strValue, I)
            
        Case 4
            lngValue = vValue
            Gf_SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lngValue, 4)
            
    End Select

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Pgm_Authority
'   2.Name         : Program ID Authority Check
'   3.Input  Value : sPgmID String
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .10
'   7.Modify Date  :
'   8.Comment      : Program ID Authority Check
'---------------------------------------------------------------------------------------
'Public Function Gf_Pgm_Authority(sPgmId As String, Optional bErpRunChk As Boolean = False) As String
'
'On Error GoTo Pgm_Authority_Error
'
'    Dim sQuery As String
'    Dim sErpRun As String
'    Dim sSec_Authority As String * 1
'    Dim sPgm_Authority As String * 4
'
'    If sUserID = "1JS6001" Or sUserID = "1JS6002" Or sUserID = "1JS6003" Or sUserID = "1JS6005" Then
'        Gf_Pgm_Authority = "1111"
'    Else
'        'Program ID Check
'        sQuery = "SELECT PGM_SECURITY FROM ZP_PGMID WHERE PGMID = '" + sPgmId + "'"
'        sSec_Authority = Gf_CodeFind(M_CN1, sQuery)
'
'        'Authority Check
'        sQuery = "SELECT INQ||INS||UPD||DEL FROM ZP_AUTHORITY WHERE EMP_ID = '" + sUserID + "' AND PGMID = '" + sPgmId + "' "
'        sPgm_Authority = Gf_CodeFind(M_CN1, sQuery)
'
'        If sSec_Authority = "1" Then                        'Inquiry Security Check
'            If Trim(sPgm_Authority) <> "" Then
'                Gf_Pgm_Authority = Trim(sPgm_Authority)
'            Else
'                Gf_Pgm_Authority = "0000"
'            End If
'        Else
'            If Trim(sPgm_Authority) <> "" Then                             'Default Inquiry Authority
'                Gf_Pgm_Authority = "1" + Right(Trim(sPgm_Authority), 3)    'Authority Check
'            Else
'                Gf_Pgm_Authority = "1000"                                  'Only Inquiry Possible
'            End If
'        End If
'
'        'ERP System Run Check
'        If bErpRunChk Then
'            sQuery = "SELECT 'Y' FROM RP_SYSTEM_RUN WHERE SYSTEM_CD = 'ERP' AND RUN_DATE <= TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS') "
'            sErpRun = Gf_CodeFind(M_CN1, sQuery)
'
'            'ERP System Run
'            If sErpRun = "Y" Then
'
'                If sSec_Authority = "1" Then                               'Inquiry Security Check
'                    If Trim(sPgm_Authority) <> "" Then
'                        Gf_Pgm_Authority = Left(Trim(sPgm_Authority), 1) + "000"
'                    Else
'                        Gf_Pgm_Authority = "0000"
'                    End If
'                Else
'                    Gf_Pgm_Authority = "1000"                              'Only Inquiry Possible
'                End If
'
'            End If
'
'        End If
'
'    End If
'
'    Exit Function
'
'Pgm_Authority_Error:
'
'    Gf_Pgm_Authority = "FAIL"
'
'End Function

Public Function Gf_Pgm_Authority(sPgmId As String, Optional bErpRunChk As Boolean = False) As String
    
On Error GoTo Pgm_Authority_Error

    Dim sQuery As String
    Dim sErpRun As String
    Dim sSec_Authority As String * 1
    Dim sPgm_Authority As String * 4
    
    If sUserID = "1JS6001" Or sUserID = "1JS6002" Or sUserID = "1JS6003" Or sUserID = "1JS6005" Then
        Gf_Pgm_Authority = "1111"
    Else
        'Program ID Check
        sQuery = "SELECT PGM_SECURITY FROM ZP_PGMID WHERE PGMID = '" + sPgmId + "'"
        sSec_Authority = Gf_CodeFind(M_CN1, sQuery)
        
        'Authority Check
'        If MainFrmType = "Old" Then ' 2012.11.09 新增  耿朝雷
'            sQuery = "SELECT INQ||INS||UPD||DEL FROM ZP_AUTHORITY WHERE EMP_ID = '" + sUserID + "' AND PGMID = '" + sPgmId + "' "
'        Else ' 2012.11.09 新增  耿朝雷
            sQuery = "SELECT PKG_ABZ_AUTHORITY.P_AUTHORITY('" + sUserID + "','" + sPgmId + "') from dual "  ' 2012.11.09 新增  耿朝雷
'        End If ' 2012.11.09 新增  耿朝雷
        sPgm_Authority = Gf_CodeFind(M_CN1, sQuery)
        
        If sSec_Authority = "1" Then                        'Inquiry Security Check
            If Trim(sPgm_Authority) <> "" Then
                Gf_Pgm_Authority = Trim(sPgm_Authority)
            Else
                Gf_Pgm_Authority = "0000"
            End If
        Else
            If Trim(sPgm_Authority) <> "" Then                             'Default Inquiry Authority
                Gf_Pgm_Authority = "1" + Right(Trim(sPgm_Authority), 3)    'Authority Check
            Else
                Gf_Pgm_Authority = "1000"                                  'Only Inquiry Possible
            End If
        End If
        
        'ERP System Run Check
        If bErpRunChk Then
            sQuery = "SELECT 'Y' FROM RP_SYSTEM_RUN WHERE SYSTEM_CD = 'ERP' AND RUN_DATE <= TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS') "
            sErpRun = Gf_CodeFind(M_CN1, sQuery)
            
            'ERP System Run
            If sErpRun = "Y" Then
            
                If sSec_Authority = "1" Then                               'Inquiry Security Check
                    If Trim(sPgm_Authority) <> "" Then
                        Gf_Pgm_Authority = Left(Trim(sPgm_Authority), 1) + "000"
                    Else
                        Gf_Pgm_Authority = "0000"
                    End If
                Else
                    Gf_Pgm_Authority = "1000"                              'Only Inquiry Possible
                End If
                
            End If
            
        End If
        
    End If
    
    Exit Function

Pgm_Authority_Error:

    Gf_Pgm_Authority = "FAIL"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Mc_Authority
'   2.Name         : Gf_Mc_Authority
'   3.Input  Value : sAuthority String, Mc Collection, {Sc Collection}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .12
'   7.Modify Date  :
'   8.Comment      : Mc Insert, Modify Authority
'---------------------------------------------------------------------------------------
Public Function Gf_Mc_Authority(sAuthority As String, MC As Collection, Optional Sc As Collection) As Boolean
    
On Error GoTo Mc_Authority_Error

    Dim iCount As Integer
    Dim sProcess As Boolean
    
    'FormType "Master", "Hsheet", "PopMaster"
    
    Select Case Mid(sAuthority, 2, 3)       'Insert, Update, Delete
    
        Case "000"      'No Authority
            Gf_Mc_Authority = False
            
        Case "001"      'Delete Authority
            
            sProcess = False
            
            If Not Sc Is Nothing Then
                
                For iCount = 1 To Sc("Spread").MaxRows
                    Sc("Spread").Col = 0: Sc("Spread").Row = iCount
                    If Sc("Spread").Text = "Input" Or Sc("Spread").Text = "Delete" Then
                        sProcess = True
                        Exit For
                    End If
                Next iCount
                
            End If
            
            If sProcess Then Gf_Mc_Authority = True: Exit Function
            
            If MC("pControl").Item(1).Enabled Then
                Call Gp_MsgBoxDisplay("权限不足  It is no input authority", "I")
                Gf_Mc_Authority = False
            Else
                Gf_Mc_Authority = True
            End If
            
        Case "010"      'Update Authority
        
            If MC("pControl").Item(1).Enabled Then
                Call Gp_MsgBoxDisplay("权限不足  It is no input authority", "I")
                Gf_Mc_Authority = False
            Else
                Gf_Mc_Authority = True
            End If
            
        Case "011"      'Update, Delete Authority
        
            If MC("pControl").Item(1).Enabled Then
                Call Gp_MsgBoxDisplay("权限不足  It is no input authority", "I")
                Gf_Mc_Authority = False
            Else
                Gf_Mc_Authority = True
            End If
            
        Case "100"      'Insert Authority
        
            If MC("pControl").Item(1).Enabled = False Then
            
                sProcess = False
                
                If Not Sc Is Nothing Then
                    
                    For iCount = 1 To Sc("Spread").MaxRows
                        Sc("Spread").Col = 0: Sc("Spread").Row = iCount
                        If Sc("Spread").Text = "Input" Or Sc("Spread").Text = "Delete" Then
                            sProcess = True
                            Exit For
                        End If
                    Next iCount
                    
                End If
                
                If sProcess Then Gf_Mc_Authority = True: Exit Function
            
                Call Gp_MsgBoxDisplay("权限不足  It is no update authority", "I")
                Gf_Mc_Authority = False
            Else
                Gf_Mc_Authority = True
            End If
        
        Case "101"      'Insert, Delete Authority
        
            sProcess = False
            
            If Not Sc Is Nothing Then
                
                For iCount = 1 To Sc("Spread").MaxRows
                    Sc("Spread").Col = 0: Sc("Spread").Row = iCount
                    If Sc("Spread").Text = "Input" Or Sc("Spread").Text = "Delete" Then
                        sProcess = True
                        Exit For
                    End If
                Next iCount
                
            End If
            
            If sProcess Then Gf_Mc_Authority = True: Exit Function
            
            If MC("pControl").Item(1).Enabled = False Then
                Call Gp_MsgBoxDisplay("权限不足  It is no update authority", "I")
                Gf_Mc_Authority = False
            Else
                Gf_Mc_Authority = True
            End If
        
        Case "110"      'Insert, Update Authority
            Gf_Mc_Authority = True
            
        Case "111"      'Insert, Update, Delete Authority
            Gf_Mc_Authority = True
            
    End Select
    
    Exit Function

Mc_Authority_Error:

    Gf_Mc_Authority = False
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sc_Authority
'   2.Name         : Gf_Sc_Authority
'   3.Input  Value : sAuthority String, iType String
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .12
'   7.Modify Date  :
'   8.Comment      : Sc Insert, Modify Authority
'---------------------------------------------------------------------------------------
Public Function Gf_Sc_Authority(sAuthority As String, iType As String) As Boolean
    
On Error GoTo Sc_Authority_Error

    'FormType "Sheet", "Msheet", "PopSheet", "Hsheet"
    If iType = "I" Then
    
        'Insert Authority Check
        If Mid(sAuthority, 2, 1) = "0" Then
            Gf_Sc_Authority = False
        Else
           Gf_Sc_Authority = True
        End If
        
    Else
        'Update Authority Check
        If Mid(sAuthority, 3, 1) = "0" Then
            Gf_Sc_Authority = False
        Else
           Gf_Sc_Authority = True
        End If
    End If
        
    Exit Function

Sc_Authority_Error:

    Gf_Sc_Authority = False
    
End Function

Public Sub OrderStausProcess()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
        
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACB3020P (?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
        
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub


Public Function CJRound(dVal As Double, iOpt As Integer, iUnit As Integer) As Double
    Dim I   As Integer
    Dim t_s As String
    Dim t_u As Double
    
    'ex) 123.4 ==> iOpt = 2(Return:123), 3(Return:124), 2(Return:123)
    'ex) 123.5 ==> iOpt = 2(Return:124), 3(Return:124), 2(Return:123)
    
    On Error Resume Next
    
    t_u = 10 ^ (iUnit * -1)
    Select Case iOpt
        Case 2:     t_u = Int((dVal + (t_u / 2)) / t_u) * t_u
        Case 3:     t_u = Int((dVal + (t_u - 10 ^ (-1 * (15 - Len(Format(Int(dVal))))))) / t_u) * t_u
        Case 4:     t_u = Int(dVal / t_u) * t_u
        Case Else:  t_u = dVal
    End Select
    
    If iUnit > 0 Then
        t_s = "0."
        For I = 1 To iUnit
            t_s = t_s + "0"
        Next I
        CJRound = Val(Format(t_u, t_s))
    Else
        CJRound = t_u
    End If
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_ErpSystem_Chk
'   2.Name         : Erp System Run Check
'   3.Input  Value :
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2008. 08 .20
'   7.Modify Date  :
'   8.Comment      : ERP System Run Check
'---------------------------------------------------------------------------------------
Public Function Gf_ErpSystem_Chk() As Boolean
    
On Error GoTo Gf_ErpSystem_Chk_Error

    Dim sQuery As String
    Dim sErpRun As String

    Gf_ErpSystem_Chk = False
    
    'ERP System Run Check
    sQuery = "SELECT 'Y' FROM RP_SYSTEM_RUN WHERE SYSTEM_CD = 'ERP' AND RUN_DATE <= TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS') "
    sErpRun = Gf_CodeFind(M_CN1, sQuery)
    
    'ERP System Run
    If sErpRun = "Y" Then
        Gf_ErpSystem_Chk = True
    End If
        
    Exit Function

Gf_ErpSystem_Chk_Error:

    Gf_ErpSystem_Chk = False

End Function
