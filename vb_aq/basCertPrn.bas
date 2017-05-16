Attribute VB_Name = "basCertPrn"
'--------------------------------------------------------------------------------------------------------
'--                                    ��Ʒ����֤����ר�ÿ���                                          --
'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn
'   2.Name         : Quality certificate public functions and subs
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      :
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn
'   2.Name         : parameters declare
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      :
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------

Option Explicit

Private xlApp       As Object   'Execel object
Private xlSheet     As Object   'Execel Sheet object
Dim Report_KND As String

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - GetPonoLot
'   2.Name         : Get certificate print date
'   3.Input  Value : sISP_SHP_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .16
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function GetPoNoLot(ByVal sISP_SHP_NO As String) As String
    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    sQuery = "Select PO_NO From QP_CERT_LOT Where ISP_SHP_NO = " + "'" + sISP_SHP_NO + "'"
    
    Set AdoRs = New adodb.Recordset
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If AdoRs.EOF Then
        GetPoNoLot = "N"
    ElseIf IsNull(AdoRs.Fields(0)) Then
        GetPoNoLot = "N"
    Else
        GetPoNoLot = AdoRs.Fields(0)
    End If
    
    AdoRs.Close
    
    Set AdoRs = Nothing
    
End Function
'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - GetCertSTD
'   2.Name         : Get certificate's Stand & Year can print CE title
'   3.Input  Value : sCERT_NO
'   4.Return Value : Boolean
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .16
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function GetCertSTD(ByVal sCert_No As String) As Boolean
    Dim sQuery      As String
    Dim AdoRs       As adodb.Recordset
    Dim iLOC        As Integer
    Dim sSTDSPEC    As String
    Dim sYEAR       As String
    Dim sSTLGRD     As String
    
    
    sQuery = "SELECT B.STDSPEC,B.STDSPEC_YY,C.STDSPEC_STLGRD "
    sQuery = sQuery + " FROM   QP_CERT_HEAD A , BP_ORDER_ITEM B ,QP_STD_HEAD C"
    sQuery = sQuery + " Where  B.ORD_NO     = A.ORD_NO "
    sQuery = sQuery + " AND    B.ORD_ITEM   = A.ORD_ITEM "
    sQuery = sQuery + " AND    C.STDSPEC    = B.STDSPEC "
    sQuery = sQuery + " AND    C.STDSPEC_YY = B.STDSPEC_YY "
    sQuery = sQuery + " AND    A.CERT_NO  = '"
    sQuery = sQuery + Trim(sCert_No) + "'"
    
    Set AdoRs = New adodb.Recordset
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If AdoRs.EOF Then
        GetCertSTD = False
    Else
        
        sSTDSPEC = AdoRs.Fields(0).Value
        sYEAR = AdoRs.Fields(1).Value
        sSTLGRD = AdoRs.Fields(2).Value
        If InStr(1, sSTDSPEC, "EN") > 0 And InStr(1, sSTDSPEC, "10025") > 0 Then
'        iLOC = InStr(1, sStdspec, "EN 10025")
         iLOC = 1
       Else
         iLOC = 0
       End If
       
        
        If iLOC = 0 Then
            GetCertSTD = False
        Else
            If sYEAR = "2004" And (Trim(sSTLGRD) <> "S355J2W" And Trim(sSTLGRD) <> "S355J0W" And Trim(sSTLGRD) <> "S355K2W") Then
                GetCertSTD = True
            Else
                GetCertSTD = False
            End If
        End If
    
    End If
    
    AdoRs.Close
    
    Set AdoRs = Nothing
    
    
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - GetPrintDate
'   2.Name         : Get certificate print date
'   3.Input  Value :
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function GetPrintDate() As String
    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    GetPrintDate = AdoRs.Fields(0)
    
    AdoRs.Close
    Set AdoRs = Nothing
    
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - Cert_type_check
'   2.Name         : Checked certificate's print function by certificate's type
'   3.Input  Value : sFlage , sCertNO ,iSave_State , sSave_Path
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Public
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Public Function Cert_type_check(ByVal sFlage As String, ByVal sCertNo As String, ByVal iSave_State As Integer, ByVal sSave_Path As String) As String
    
    Select Case sFlage
        
        Case "C"                        '�����Ʒ
            Cert_type_check = funGetQuery_C(sCertNo, iSave_State, sSave_Path)
        Case "S"                        '�����ʱ���
            Cert_type_check = funGetQuery_S(sCertNo, iSave_State, sSave_Path)
        Case "T"                        '����˵����
            Report_KND = "T"
            Cert_type_check = funGetQuery_S(sCertNo, iSave_State, sSave_Path)
            Report_KND = ""
        Case "P"                        '���߸�
            Cert_type_check = funGetQuery_P(sCertNo, iSave_State, sSave_Path)
        Case "B"                        '����
            Cert_type_check = funGetQuery_B(sCertNo, iSave_State, sSave_Path)
        Case Else                       '��������
            Cert_type_check = "Err Type"
    End Select
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - Cert_Save
'   2.Name         : Save certificate
'   3.Input  Value : oNewBook , sCertNo , sPageNO , iSave_State ,sSave_Path
'   4.Return Value : Integer ( 0,1,2 )
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function Cert_Save(ByVal oNewBook As Object, ByVal sCertNo As String, ByVal sPageNO As String, ByVal iSave_State As Integer, ByVal sSave_Path As String) As Integer
    
    Dim Save_FileName  As String
    Dim Save_Path       As String
    Dim Save_State      As Integer
    
    Save_State = iSave_State
    
    If Trim(sSave_Path) = "" Or Len(Trim(sSave_Path)) = 0 Then
        Save_Path = App.Path
    Else
        Save_Path = sSave_Path
    End If
    
    If Save_State = 0 Then
        Save_Path = "N"
        Save_FileName = "N"
    Else
        If InStr(4, Save_Path, "\") > 0 Then
            Save_FileName = Trim(Save_Path) + "\" + Trim(sCertNo) + "-" + Trim(sPageNO) + ".xls"
        Else
            If Len(Save_Path) <> 3 And InStr(3, Save_Path, "\") > 0 Then
                Save_FileName = Trim(Save_Path) + "\" + Trim(sCertNo) + "-" + Trim(sPageNO) + ".xls"
            Else
                Save_FileName = Trim(Save_Path) + Trim(sCertNo) + "-" + Trim(sPageNO) + ".xls"
            End If
        End If
    End If
    
    If Save_FileName = "N" Then
        Cert_Save = 0
        Exit Function
    Else
        Cert_Save = Save_State
        oNewBook.SaveAs FileName:=Save_FileName
        Exit Function
    End If
End Function
'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - GetChem_Rslt_SQL
'   2.Name         : Get SQL for quering certificate chemical result
'   3.Input  Value : sCert_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 11 .26
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Function GetChem_Rslt_SQL(ByVal sCert_No As String) As String
    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim sMesg As String
    
    
    Dim adoCmd As adodb.Command
    Screen.MousePointer = vbHourglass
    

    OutParam(1, 1) = "arg_SQL"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 8000
    

    sQuery = "{call GP_GET_QLTYCHEM_CERT_SQL('" + sCert_No + "',?)}"
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    
    'Process Error Check
    If adoCmd("arg_SQL") = "NN" Then
        ret_Result_ErrMsg = "��ȡʧ�ܣ������ʱ�������"
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        GetChem_Rslt_SQL = "NN"
        Set adoCmd = Nothing
        Exit Function
        
    End If
    
    GetChem_Rslt_SQL = adoCmd("arg_SQL")
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - funGetQuery_B
'   2.Name         : Slab certificate
'   3.Input  Value : sCertNo , iSave_State ,sSave_Path
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function funGetQuery_B(ByVal sCertNo As String, ByVal iSave_State As Integer, ByVal sSave_Path As String) As String
        
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
    Dim AdoRs As adodb.Recordset
           
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT CERT_NO , PROD_SPEC_NO , STDSPEC_NAME "
    sQuery = sQuery + ", DECODE(TRIM(GF_ENDCUSTER_FIND(SHIP_ISP_NO)),'',GF_CUST_NAME(CUST_CD,'')) AS ENDCUSTER_NAME "
    sQuery = sQuery + ",SHIP_ISP_NO, GF_PONO_FIND(ORD_NO) AS PONO"
    sQuery = sQuery + " FROM QP_CERT_HEAD WHERE CERT_NO = '" & sCertNo & "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funGetQuery_B = "Err Database"
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = "SELECT CERT_NO ,PROD_NO ,STLGRD ,PROD_SIZE ,PRDT_QNTY , PRDT_WGT "
    sQuery = sQuery + ",DECODE(C_RST,NULL,0,C_RST)  ,DECODE(MN_RST,NULL,0,MN_RST), DECODE(P_RST,NULL,0,P_RST)"
    sQuery = sQuery + ",DECODE(S_RST,NULL,0,S_RST), DECODE(SI_RST,NULL,0,SI_RST),DECODE(CU_RST,NULL,0,CU_RST)"
    sQuery = sQuery + ",DECODE(NI_RST,NULL,0,NI_RST),DECODE(CR_RST,NULL,0,CR_RST),DECODE(MO_RST,NULL,0,MO_RST)"
    sQuery = sQuery + ",DECODE(V_RST,NULL,0,V_RST),DECODE(TI_RST,NULL,0,TI_RST),DECODE(NB_RST,NULL,0,NB_RST)"
    sQuery = sQuery + ",DECODE(AL_RST,NULL,0,AL_RST),DECODE(CEQ_RST,NULL,0,CEQ_RST) "
    sQuery = sQuery + ",GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Alt'),GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'N')"
    sQuery = sQuery + ",GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Pcm'),GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'B')"
    sQuery = sQuery + ",GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Sn'),GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Ca')"
    sQuery = sQuery + ",GF_MARK_STRING(PROD_NO)"
    sQuery = sQuery + "  FROM QP_CERT_DETAIL  WHERE CERT_NO  = '" & sCertNo & "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funGetQuery_B = "Err Database"
        Exit Function
    End If
    arrRecords2 = AdoRs.GetRows
    AdoRs.Close
       
    Set AdoRs = Nothing
    
    funGetQuery_B = MillSheetPrint_B(iSave_State, sSave_Path, arrRecords1, arrRecords2)
       
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_B
'   2.Name         : Slab certificate print(detail table)
'   3.Input  Value : iSave_State ,sSave_Path ,arrRecords1 ,arrRecords2
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function MillSheetPrint_B(ByVal iSave_State As Integer, ByVal sSave_Path As String, ByVal arrRecords1 As Variant, ByVal arrRecords2 As Variant) As String
    Dim RowCNT          As Long
    Dim PrtCnt          As Long
    Dim LneCnt          As Long
    Dim pAry11()        As String                   'PROD_NO AND STLGRD
    Dim pAry12()        As String                   'PROD_SIZE
    Dim pAry13()        As String                   'PIECE
    Dim pAry14()        As String                   'WEIGHT
    Dim pAry15()        As String                   'CHEM_RSLT
    
    Dim lSumQNTY        As Long                     'PAGE PIECE COUNT
    Dim dSumWGT         As Double                   'PAGE WEIGHT COUNT
    Dim lSumQNTY_T      As Long                     'ALL PIECE COUNT
    Dim dSumWGT_T       As Double                   'ALL WEIGHT COUNT
    
    Dim STLGRD          As String
    Dim Save_State      As Integer
    Dim Save_Path       As String
    Dim Cert_No         As String
    Dim Page_no         As Integer
    
    
    Save_State = iSave_State
    Save_Path = sSave_Path
    Cert_No = arrRecords2(0, 0)
    
    If IsEmpty(arrRecords1) Or IsEmpty(arrRecords2) Then
        MillSheetPrint_B = "Err Data"
        Exit Function
    End If
    
    RowCNT = UBound(arrRecords2, 2)
    
    PrtCnt = -1
    LneCnt = 0
    lSumQNTY = 0
    dSumWGT = 0
    lSumQNTY_T = 0
    dSumWGT_T = 0
    
    ReDim pAry11(1 To 20, 1 To 2)
    ReDim pAry12(1 To 20, 1 To 1)
    ReDim pAry13(1 To 20, 1 To 1)
    ReDim pAry14(1 To 20, 1 To 1)
    ReDim pAry15(1 To 20, 1 To 21)
    
    Do

        LneCnt = LneCnt + 1
        PrtCnt = PrtCnt + 1

        pAry11(LneCnt, 1) = arrRecords2(1, PrtCnt) & ""                 ' PROD_NO
            
        pAry11(LneCnt, 2) = arrRecords2(2, PrtCnt) & ""                 ' STLGRD
        
        STLGRD = arrRecords2(2, PrtCnt)
        
        pAry12(LneCnt, 1) = arrRecords2(3, PrtCnt) & ""                 ' PROD_SIZE
        pAry13(LneCnt, 1) = arrRecords2(4, PrtCnt) & ""                 ' QNTY
        pAry14(LneCnt, 1) = arrRecords2(5, PrtCnt) & ""                 ' WGT
        
        pAry15(LneCnt, 1) = IIf(Val(arrRecords2(6, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(6, PrtCnt) & "") * 100)                   ' C_RST
        
        pAry15(LneCnt, 2) = IIf(Val(arrRecords2(7, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(7, PrtCnt) & "") * 100)                   ' MN_RST
        
        pAry15(LneCnt, 3) = IIf(Val(arrRecords2(8, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(8, PrtCnt) & "") * 1000)                   ' P_RST
        
        pAry15(LneCnt, 4) = IIf(Val(arrRecords2(9, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(9, PrtCnt) & "") * 1000)                  ' S_RST
        
        pAry15(LneCnt, 5) = IIf(Val(arrRecords2(10, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(10, PrtCnt) & "") * 100)                  ' SI_RST
        
        pAry15(LneCnt, 6) = IIf(Val(arrRecords2(11, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(11, PrtCnt) & "") * 100)                 ' CU_RST
        
        pAry15(LneCnt, 7) = IIf(Val(arrRecords2(12, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(12, PrtCnt) & "") * 100)                 ' NI_RST
        
        pAry15(LneCnt, 8) = IIf(Val(arrRecords2(13, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(13, PrtCnt) & "") * 100)                 ' CR_RST
        
        pAry15(LneCnt, 9) = IIf(Val(arrRecords2(14, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(14, PrtCnt) & "") * 1000)                 ' MO_RST
        
        pAry15(LneCnt, 10) = IIf(Val(arrRecords2(15, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(15, PrtCnt) & "") * 1000)                 ' V_RST
        
        pAry15(LneCnt, 11) = IIf(Val(arrRecords2(16, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(16, PrtCnt) & "") * 1000)                 ' TI_RST
        
        pAry15(LneCnt, 12) = IIf(Val(arrRecords2(17, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(17, PrtCnt) & "") * 1000)                 ' NB_RST
        
        pAry15(LneCnt, 13) = IIf(Val(arrRecords2(20, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(20, PrtCnt) & "") * 1000)                 ' ALT_RST
        
'        pAry15(LneCnt, 14) = IIf(Val(arrRecords2(18, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(18, PrtCnt) & "") * 1000)                 ' AL_RST
        pAry15(LneCnt, 14) = ""
        
        pAry15(LneCnt, 15) = IIf(Val(arrRecords2(23, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(23, PrtCnt) & "") * 1000)                 ' B_RST
        
        pAry15(LneCnt, 16) = IIf(Val(arrRecords2(24, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(24, PrtCnt) & "") * 1000)                 ' SN_RST
'Orig.
'        pAry15(LneCnt, 16) = IIf(Val(arrRecords2(25, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(25, PrtCnt) & "") * 1000)                 ' CA_RST
'
'        pAry15(LneCnt, 16) = IIf(Val(arrRecords2(21, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(21, PrtCnt) & ""))                        ' N_RST
'
'        pAry15(LneCnt, 16) = IIf(Val(arrRecords2(19, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(19, PrtCnt) & "") * 100)                  ' CEQ_RST
'
'        pAry15(LneCnt, 16) = IIf(Val(arrRecords2(22, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(22, PrtCnt) & "") * 100)                  ' PCM_RST

        pAry15(LneCnt, 17) = IIf(Val(arrRecords2(25, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(25, PrtCnt) & "") * 1000)                 ' CA_RST

        pAry15(LneCnt, 18) = IIf(Val(arrRecords2(21, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(21, PrtCnt) & ""))                        ' N_RST

        pAry15(LneCnt, 19) = IIf(Val(arrRecords2(19, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(19, PrtCnt) & "") * 100)                  ' CEQ_RST

        pAry15(LneCnt, 20) = IIf(Val(arrRecords2(22, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(22, PrtCnt) & "") * 100)                  ' PCM_RST
        pAry15(LneCnt, 21) = arrRecords2(26, PrtCnt) & ""              ' osm_no
        
       
       
        lSumQNTY = lSumQNTY + arrRecords2(4, PrtCnt)                    'PAGE SUM QUANTITY
        dSumWGT = dSumWGT + arrRecords2(5, PrtCnt)                      'PAGE SUM WEIGHT
        
        lSumQNTY_T = lSumQNTY_T + arrRecords2(4, PrtCnt)                'TOTAL SUM QUANTITY
        dSumWGT_T = dSumWGT_T + arrRecords2(5, PrtCnt)                  'TOTAL SUM WEIGHT
       
        If LneCnt = 20 Then
           
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
        
            xlApp.Workbooks.Open (App.Path & "\AQD070C.xls")
            Set xlSheet = xlApp.Worksheets("Sheet1")
            
            Call MillSheetPrint_B_Head(arrRecords1)
                        
            xlSheet.Range("B36").Value = lSumQNTY & ""      'SUM_CNT
            xlSheet.Range("C36").Value = dSumWGT & ""        'SUM_WGT
            xlSheet.Range("AC36").Value = Round((PrtCnt + 10) / 20, 0)    'PAGE NO.
            Page_no = Round((PrtCnt + 10) / 20, 0)
            
            If PrtCnt = RowCNT Then
               xlSheet.Range("X35").Value = lSumQNTY_T & " Piece"        'SUM_CNT
               xlSheet.Range("X36").Value = dSumWGT_T & " ton"        'SUM_WGT
            End If
            
            
            xlSheet.Range("B14:C33").Value = pAry11
            xlSheet.Range("D14:D33").Value = pAry12
            xlSheet.Range("F14:F33").Value = pAry13
            xlSheet.Range("G14:G33").Value = pAry14
            xlSheet.Range("J14:AD33").Value = pAry15
            
            Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, iSave_State, sSave_Path)
            If Save_State = 0 Or Save_State = 1 Then
            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            End If
            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit

            LneCnt = 0
            lSumQNTY = 0
            dSumWGT = 0
            
            ReDim pAry11(1 To 20, 1 To 2)
            ReDim pAry12(1 To 20, 1 To 1)
            ReDim pAry13(1 To 20, 1 To 1)
            ReDim pAry14(1 To 20, 1 To 1)
            ReDim pAry15(1 To 20, 1 To 21)
            
        End If

    Loop Until PrtCnt = RowCNT
    
    If LneCnt <> 0 Then
    
        
        Set xlApp = GetObject("", "Excel.Application")
        If Err.Number = 429 Then
            Set xlApp = CreateObject("", "Excel.Application")
        End If
    
        xlApp.Workbooks.Open (App.Path & "\AQD070C.xls")
        Set xlSheet = xlApp.Worksheets("Sheet1")
        
            Call MillSheetPrint_B_Head(arrRecords1)
        
            xlSheet.Range("B36").Value = lSumQNTY & ""      'SUM_CNT
            xlSheet.Range("C36").Value = dSumWGT & ""        'SUM_WGT
            xlSheet.Range("AC36").Value = Round((PrtCnt + 10) / 20, 0)    'PAGE NO.
            Page_no = Round((PrtCnt + 10) / 20, 0)
            
            If PrtCnt = RowCNT Then
               xlSheet.Range("X35").Value = lSumQNTY_T & " Piece"        'SUM_CNT
               xlSheet.Range("X36").Value = dSumWGT_T & " ton"        'SUM_WGT
            End If
            
            
            xlSheet.Range("B14:C33").Value = pAry11
            xlSheet.Range("D14:D33").Value = pAry12
            xlSheet.Range("F14:F33").Value = pAry13
            xlSheet.Range("G14:G33").Value = pAry14
            xlSheet.Range("J14:AD33").Value = pAry15
            
            Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, iSave_State, sSave_Path)
            If Save_State = 0 Or Save_State = 1 Then
            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            End If
            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit
            
    End If
        
    Set xlApp = Nothing
    
    Exit Function
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_B_Head
'   2.Name         : Slab certificate print(Head table)
'   3.Input  Value : arrRecords1
'   4.Return Value :
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Sub MillSheetPrint_B_Head(arrRecords1 As Variant)
    Dim sDate As String
    Dim sPONO As String
    
    sPONO = GetPoNoLot(arrRecords1(4, 0))
    
    sDate = GetPrintDate()
    
    xlSheet.Range("C2").Value = arrRecords1(0, 0) & ""         'CERT_NO
    xlSheet.Range("C4").Value = arrRecords1(1, 0) & ""         'PROD_SPEC_NO
    xlSheet.Range("C6").Value = arrRecords1(2, 0) & ""         'STDSPEC_NAME
    xlSheet.Range("V4").Value = arrRecords1(4, 0) & ""         'SHIP_ISP_NO
    If sPONO = "N" Then
        xlSheet.Range("V6").Value = arrRecords1(5, 0) & ""         'PONO
    Else
        xlSheet.Range("V6").Value = sPONO & ""         'PONO
    End If
    xlSheet.Range("V8").Value = arrRecords1(3, 0) & ""         'CUSTER NAME
    xlSheet.Range("N34").Value = sUserName & ""                 'TEST_EMP
    xlSheet.Range("AC35").Value = sDate                         'PRINT DATE

End Sub

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - funGetQuery_C
'   2.Name         : Conventionality certificate
'   3.Input  Value : sCertNo , iSave_State ,sSave_Path
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function funGetQuery_C(sCertNo As String, iSave_State As Integer, sSave_Path As String) As String
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
    Dim arrRecords3 As Variant
    Dim AdoRs As adodb.Recordset
    Dim sPROD_CD As String
    Dim sTable_PROD As String
    Dim sFieldName_NO As String
    
    sPROD_CD = Mid(sCertNo, 1, 2)
    If UCase(sPROD_CD) = "HC" Then
        sTable_PROD = "GP_COIL"
        sFieldName_NO = "COIL_NO"
    Else
        sTable_PROD = "GP_PLATE"
        sFieldName_NO = "PLATE_NO"
    End If
   
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT CERT_NO , PROD_NAME ,STDSPEC_NAME , PROD_SPEC_NO , DECODE(TRIM(GF_ENDCUSTER_FIND(SHIP_ISP_NO)),'',GF_CUST_NAME(CUST_CD,''),GF_ENDCUSTER_FIND(SHIP_ISP_NO)) "
    sQuery = sQuery + ",COND_SUPPLY , PROD_SIZE,IMPACT_SMP_SIZE , QLTY_REC_NO , GF_PONO_FIND(ORD_NO), SHIP_ISP_NO , TRAIN_LINE_NAME"
    sQuery = sQuery + ",DEST_DETAIL , CERT_RPT_DATE , BEND_DIA , GF_EMPNAMEFIND(TEST_EMP) AS TEST_EMP , GF_EMPNAMEFIND(SHP_EMP) AS SHP_EMP"
    sQuery = sQuery + ",DECODE(Gf_ComnNameFind('Q0046',UST_FL) ,'ASTM A 435 / ASME SA-435','ASTM A435 / A435M-90','JB4730 J11'"
    sQuery = sQuery + ",'JB4730-94 ��','JB4730 J21','JB4730-94 ��','GB/T 2970 K11','GB/T 2970 ��','GB/T 2970 K21','GB/T 2970 ��','NO UST',' '"
    sQuery = sQuery + ",Gf_ComnNameFind('Q0046',UST_FL))"
    sQuery = sQuery + ",AQD0060C.F_SUM_CNT(CERT_NO) AS SUM_CNT, AQD0060C.F_SUM_WGT(CERT_NO) AS SUM_WGT, PROD_DGR,GF_STDSPEC_NAME_ENG(STDSPEC_STLGRD)"
    sQuery = sQuery + " FROM QP_CERT_HEAD WHERE CERT_NO = '" & sCertNo & "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_C = "Err DataBase"
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = GetChem_Rslt_SQL(sCertNo)
    
    If sQuery = "NN" Then
        sQuery = "SELECT CERT_NO ,GF_MARKING_NO(A.PROD_NO) ,GF_STLGRD_DETAIL(A.STLGRD) ,DECODE(SUBSTR(A.PROD_SIZE,-1),'L', B.ORD_THK||'*'||B.ORD_WID||'*'||B.LEN,A.PROD_SIZE )"
        sQuery = sQuery + ", PRDT_QNTY , PRDT_WGT"
        sQuery = sQuery + ", DECODE(C_RST,NULL,0,C_RST)  ,DECODE(MN_RST,NULL,0,MN_RST), DECODE(P_RST,NULL,0,P_RST),      DECODE(S_RST,NULL,0,S_RST) "
        sQuery = sQuery + ", DECODE(SI_RST,NULL,0,SI_RST),DECODE(NB_RST,NULL,0,NB_RST), GF_AQD0060C_FIND('Alt',PROD_NO,CERT_NO) ,DECODE(MO_RST,NULL,0,MO_RST) "
        sQuery = sQuery + ", DECODE(CU_RST,NULL,0,CU_RST),DECODE(NI_RST,NULL,0,NI_RST), DECODE(CR_RST,NULL,0,CR_RST)    ,DECODE(V_RST,NULL,0,V_RST) "
        sQuery = sQuery + ", DECODE(TI_RST,NULL,0,TI_RST),GF_AQD0060C_FIND('N',PROD_NO,CERT_NO),DECODE(CEQ_RST,NULL,0,CEQ_RST) , GF_AQD0060C_FIND('Pcm',PROD_NO,CERT_NO)"
        sQuery = sQuery + ",GF_AQD0060C_ADD(CHEM_COMP_CD1),DECODE(CHEM_COMP_CD1,NULL,0,GF_AQD0060C_FIND(CHEM_COMP_CD1,PROD_NO,CERT_NO))"
        sQuery = sQuery + ",GF_AQD0060C_ADD(CHEM_COMP_CD2),DECODE(CHEM_COMP_CD2,NULL,0,GF_AQD0060C_FIND(CHEM_COMP_CD2,PROD_NO,CERT_NO))"
        sQuery = sQuery + ",GF_AQD0060C_ADD(CHEM_COMP_CD3),DECODE(CHEM_COMP_CD3,NULL,0,GF_AQD0060C_FIND(CHEM_COMP_CD3,PROD_NO,CERT_NO))"
        sQuery = sQuery + ",GF_AQD0060C_ADD(CHEM_COMP_CD4),DECODE(CHEM_COMP_CD4,NULL,0,GF_AQD0060C_FIND(CHEM_COMP_CD4,PROD_NO,CERT_NO))"
        sQuery = sQuery + ",GF_AQD0060C_ADD(CHEM_COMP_CD5),DECODE(CHEM_COMP_CD5,NULL,0,GF_AQD0060C_FIND(CHEM_COMP_CD5,PROD_NO,CERT_NO))"
        sQuery = sQuery + ",GF_AQD0060C_FIND('O',PROD_NO,CERT_NO),GF_AQD0060C_FIND('H',PROD_NO,CERT_NO),GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'B'),"
        sQuery = sQuery + " GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Sn')"
        sQuery = sQuery + "  FROM QP_CERT_DETAIL  A ," + sTable_PROD + " B WHERE CERT_NO  = '" & sCertNo & "' AND A.PROD_NO=B." + sFieldName_NO + " ORDER BY A.PROD_NO"
    End If
    
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_C = "Err DataBase"
        Exit Function
    End If
    arrRecords2 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = "SELECT CERT_NO ,GF_MARKING_NO(PROD_NO) "
    sQuery = sQuery + ", DECODE(YP_RST,NULL,0,YP_RST) , DECODE(TS_RST,NULL,0,TS_RST) , DECODE(EL_RST,NULL,0,EL_RST) , DECODE(BEND_RST,'Y','OK','-')  "
    sQuery = sQuery + ", DECODE (UST_GRD_RST,NULL,'','OK') , DECODE(IMPACT_TMP,NULL,0,IMPACT_TMP) , DECODE(IMPACT_RST1,NULL,0,IMPACT_RST1)"
    sQuery = sQuery + ", DECODE(IMPACT_RST2,NULL,0,IMPACT_RST2) , DECODE(IMPACT_RST3,NULL,0,IMPACT_RST3),  DECODE(IMPACT_RST4,NULL,0,IMPACT_RST4)"
    sQuery = sQuery + ", DECODE(IMPACT_RST5,NULL,0,IMPACT_RST5) , DECODE(IMPACT_RST6,NULL,0,IMPACT_RST6), DECODE(IMPACT_RST_AVE,NULL,0,IMPACT_RST_AVE) "
    sQuery = sQuery + ", DECODE(TIM_IMPACT_TMP,NULL,0,TIM_IMPACT_TMP) , DECODE(TIM_IMPACT_RST1,NULL,0,TIM_IMPACT_RST1),DECODE(TIM_IMPACT_RST2,NULL,0,TIM_IMPACT_RST2)"
    sQuery = sQuery + ", DECODE(TIM_IMPACT_RST3,NULL,0,TIM_IMPACT_RST3),DECODE(TIM_IMPACT_RST4,NULL,0,TIM_IMPACT_RST4),DECODE(TIM_IMPACT_RST5,NULL,0,TIM_IMPACT_RST5)"
    sQuery = sQuery + ", DECODE(TIM_IMPACT_RST6,NULL,0,TIM_IMPACT_RST6),DECODE(TIM_IMPACT_RST_AVE,NULL,0,TIM_IMPACT_RST_AVE),DECODE(RA_RST,NULL,0,RA_RST)"
    sQuery = sQuery + ", DECODE(RA_RST2,NULL,0,RA_RST2),DECODE(RA_RST3,NULL,0,RA_RST3),DECODE(RA_RST_AVE,NULL,0,RA_RST_AVE),DECODE(GRAIN_SIZE_RST,NULL,0,GRAIN_SIZE_RST)"
    sQuery = sQuery + ", DECODE(NON_METAL_DSC_RST,'Y','OK','N','NO','0','OK'),DECODE(YR_RST,NULL,0,YR_RST),GF_GET_HARDRSLT(PROD_NO,'PP')"
    sQuery = sQuery + "   FROM QP_CERT_DETAIL  WHERE CERT_NO  = '" & sCertNo & "'" & " ORDER BY PROD_NO"

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_C = "Err DataBase"
        Exit Function
    End If
    arrRecords3 = AdoRs.GetRows
    AdoRs.Close
    
    Set AdoRs = Nothing
       
    funGetQuery_C = MillSheetPrint_C(iSave_State, sSave_Path, arrRecords1, arrRecords2, arrRecords3)
    
End Function


'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_C
'   2.Name         : Conventionality certificate print(Detail table)
'   3.Input  Value : iSave_State ,sSave_Path ,arrRecords1 ,arrRecords2,arrRecords3
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function MillSheetPrint_C(ByVal iSave_State As Integer, ByVal sSave_Path As String, arrRecords1 As Variant, arrRecords2 As Variant, arrRecords3 As Variant) As String
    Dim RowCNT                  As Long
    Dim ColCnt                  As Long
    Dim PrtCnt                  As Long
    Dim LneCnt                  As Long
    
    Dim pAry11()                As String
    Dim pAry12()                As String
    Dim pAry13()                As String
    Dim pAry14()                As String
    Dim pAry15()                As String
    Dim pAry2()                 As String
    
    Dim lSumQNTY                As Long
    Dim dSumWGT                 As Double
    Dim lSumQNTY_T              As Long
    Dim dSumWGT_T               As Double
    
    Dim STLGRD                  As String
    Dim CHEM_ADD1               As String
    Dim CHEM_ADD2               As String
    Dim CHEM_ADD3               As String
    
    
    Dim ADD1, ADD2, ADD3, ADD4  As Integer
    Dim Save_State              As Integer
    Dim Save_Path               As String
    Dim Page_no                 As String
    Dim Cert_No                 As String
    Dim i_ID                    As Integer
    Dim s_CHEM_CD               As String
    
    
    If IsEmpty(arrRecords1) Or IsEmpty(arrRecords2) Or IsEmpty(arrRecords3) Then
        MillSheetPrint_C = "Err Data"
        Exit Function
    End If
    
    Save_State = iSave_State
    Save_Path = sSave_Path
    Cert_No = arrRecords2(0, 0)
    
    RowCNT = UBound(arrRecords2, 2)
    ColCnt = UBound(arrRecords2, 1)
    
    PrtCnt = -1
    LneCnt = 0
    lSumQNTY = 0
    dSumWGT = 0
    lSumQNTY_T = 0
    dSumWGT_T = 0
'--------------------------------------------------------------------------------------------------------------------------------------
    
    ReDim pAry11(1 To 6, 1 To 6)                                    '��Ʒ��Ϣ(��Ʒ��\�ƺ�\�ߴ�\֧��\����)
    ReDim pAry12(1 To 1, 1 To 18)                                   '�ɷ�����
    ReDim pAry13(1 To 1, 1 To 18)                                   '�ɷֱ���
    ReDim pAry14(1 To 5, 1 To 18)                                   '�ɷ�ʵ��
    ReDim pAry15(1 To 1, 1 To 2)                                    '���óɷ�����
    ReDim pAry16(1 To 1, 1 To 2)                                    '���óɷֱ���
    ReDim pAry17(1 To 5, 1 To 2)                                    '���óɷ�ʵ��
    
    ReDim pAry2(1 To 5, 1 To 23)                                    '����ʵ��(��һλ��Ʒ��)
'--------------------------------------------------------------------------------------------------------------------------------------
    
    Do

        LneCnt = LneCnt + 1
        PrtCnt = PrtCnt + 1
        
        STLGRD = arrRecords2(2, PrtCnt)
'--------------------------------------------------------------�ɷ�����----------------------------------------------------------------
        
        pAry11(LneCnt, 1) = arrRecords2(1, PrtCnt) & ""                 ' PROD_NO
        pAry11(LneCnt, 2) = arrRecords2(2, PrtCnt) & ""                 ' STLGRD
        pAry11(LneCnt, 3) = arrRecords2(3, PrtCnt) & ""                 ' PROD_SIZE
        pAry11(LneCnt, 4) = "" & ""                                     ' PROD_SIZE/""
        pAry11(LneCnt, 5) = arrRecords2(4, PrtCnt) & ""                 ' QNTY
        pAry11(LneCnt, 6) = arrRecords2(5, PrtCnt) & ""                 ' WGT
        
        For i_ID = 6 To ColCnt - 2 Step 2
            s_CHEM_CD = Trim(arrRecords2(i_ID + 1, PrtCnt))
            Select Case UCase(s_CHEM_CD)
                Case "C", "MN", "SI", "CU", "NI", "CR", "CEQ", "PCM"
                        pAry12(1, i_ID / 2 - 2) = s_CHEM_CD
                        pAry13(1, i_ID / 2 - 2) = "X100"
                        pAry14(LneCnt, i_ID / 2 - 2) = IIf(Val(arrRecords2(i_ID, PrtCnt) & "") = 0, _
                                                        "-", Val(arrRecords2(i_ID, PrtCnt) & "") * 100)
                Case "O", "N", "H"
                        pAry12(1, i_ID / 2 - 2) = s_CHEM_CD
                        pAry13(1, i_ID / 2 - 2) = "ppm"
                        pAry14(LneCnt, i_ID / 2 - 2) = IIf(Val(arrRecords2(i_ID, PrtCnt) & "") = 0, _
                                                        "-", Val(arrRecords2(i_ID, PrtCnt) & ""))
                Case Else
                        pAry12(1, i_ID / 2 - 2) = s_CHEM_CD
                        pAry13(1, i_ID / 2 - 2) = "X1000"
                        pAry14(LneCnt, i_ID / 2 - 2) = IIf(Val(arrRecords2(i_ID, PrtCnt) & "") = 0, _
                                                        "-", Val(arrRecords2(i_ID, PrtCnt) & "") * 1000)
            End Select
            If i_ID > 40 Then Exit For
        Next i_ID
        
        If ColCnt > 42 Then
            For i_ID = 42 To 44 Step 2
            s_CHEM_CD = Trim(arrRecords2(i_ID + 1, PrtCnt))
            Select Case UCase(s_CHEM_CD)
                Case "C", "MN", "SI", "CU", "NI", "CR", "CEQ", "PCM"
                        pAry12(1, i_ID / 2 - 20) = s_CHEM_CD
                        pAry13(1, i_ID / 2 - 20) = "X100"
                        pAry14(LneCnt, i_ID / 2 - 20) = IIf(Val(arrRecords2(i_ID, PrtCnt) & "") = 0, _
                                                        "-", Val(arrRecords2(i_ID, PrtCnt) & "") * 100)
                Case "O", "N", "H"
                        pAry12(1, i_ID / 2 - 20) = s_CHEM_CD
                        pAry13(1, i_ID / 2 - 20) = "ppm"
                        pAry14(LneCnt, i_ID / 2 - 20) = IIf(Val(arrRecords2(i_ID, PrtCnt) & "") = 0, _
                                                        "-", Val(arrRecords2(i_ID, PrtCnt) & ""))
                Case Else
                        pAry12(1, i_ID / 2 - 20) = s_CHEM_CD
                        pAry13(1, i_ID / 2 - 20) = "X1000"
                        pAry14(LneCnt, i_ID / 2 - 20) = IIf(Val(arrRecords2(i_ID, PrtCnt) & "") = 0, _
                                                        "-", Val(arrRecords2(i_ID, PrtCnt) & "") * 1000)
            End Select
            Next i_ID
        End If
'--------------------------------------------------------------��������----------------------------------------------------------------
        pAry2(LneCnt, 1) = arrRecords3(1, PrtCnt) & ""                                          ' PROD_NO
 
        pAry2(LneCnt, 2) = IIf(arrRecords3(2, PrtCnt) = 0, "-", arrRecords3(2, PrtCnt) & "")    ' YP_RST
        pAry2(LneCnt, 3) = IIf(arrRecords3(3, PrtCnt) = 0, "-", arrRecords3(3, PrtCnt) & "")    ' TS_RST
        pAry2(LneCnt, 4) = IIf(arrRecords3(4, PrtCnt) = 0, "-", arrRecords3(4, PrtCnt) & "")    ' EL_RST
        pAry2(LneCnt, 5) = IIf(arrRecords3(5, PrtCnt) = 0, "-", arrRecords3(5, PrtCnt) & "")    ' BEND_RST
        pAry2(LneCnt, 6) = arrRecords3(6, PrtCnt) & ""                                          ' UST_GRD_RST

        
        pAry2(LneCnt, 7) = IIf(arrRecords3(8, PrtCnt) = 0, "-", arrRecords3(7, PrtCnt) & "")    ' IMPACT_TMP
        pAry2(LneCnt, 8) = IIf(arrRecords3(8, PrtCnt) = 0, "-", arrRecords3(8, PrtCnt) & "")    ' IMPACT_RST1
        pAry2(LneCnt, 9) = IIf(arrRecords3(9, PrtCnt) = 0, "-", arrRecords3(9, PrtCnt) & "")    ' IMPACT_RST2
        pAry2(LneCnt, 10) = IIf(arrRecords3(10, PrtCnt) = 0, "-", arrRecords3(10, PrtCnt) & "") ' IMPACT_RST3
        pAry2(LneCnt, 11) = IIf(arrRecords3(14, PrtCnt) = 0, "-", arrRecords3(14, PrtCnt) & "") ' IMPACT_RST_AVE
        
        If arrRecords3(11, PrtCnt) > 0 Or arrRecords3(12, PrtCnt) > 0 Or arrRecords3(13, PrtCnt) > 0 Then
            pAry2(LneCnt, 8) = IIf(arrRecords3(14, PrtCnt) = 0, "-", arrRecords3(14, PrtCnt) & "") ' IMPACT_RST_AVE
            pAry2(LneCnt, 9) = IIf(arrRecords3(14, PrtCnt) = 0, "-", arrRecords3(14, PrtCnt) & "") ' IMPACT_RST_AVE
            pAry2(LneCnt, 10) = IIf(arrRecords3(14, PrtCnt) = 0, "-", arrRecords3(14, PrtCnt) & "") ' IMPACT_RST_AVE
        End If
        
        pAry2(LneCnt, 12) = IIf(arrRecords3(15, PrtCnt) = 0, "-", arrRecords3(15, PrtCnt) & "") ' TIM_IMPACT_TMP
        pAry2(LneCnt, 13) = IIf(arrRecords3(16, PrtCnt) = 0, "-", arrRecords3(16, PrtCnt) & "") ' TIM_IMPACT_RST_1
        pAry2(LneCnt, 14) = IIf(arrRecords3(17, PrtCnt) = 0, "-", arrRecords3(17, PrtCnt) & "") ' TIM_IMPACT_RST_2
        pAry2(LneCnt, 15) = IIf(arrRecords3(18, PrtCnt) = 0, "-", arrRecords3(18, PrtCnt) & "") ' TIM_IMPACT_RST_3
        pAry2(LneCnt, 16) = IIf(arrRecords3(22, PrtCnt) = 0, "-", arrRecords3(22, PrtCnt) & "") ' TIM_IMPACT_RST_AVE
        
        If arrRecords3(19, PrtCnt) > 0 Or arrRecords3(20, PrtCnt) > 0 Or arrRecords3(21, PrtCnt) > 0 Then
            pAry2(LneCnt, 8) = IIf(arrRecords3(22, PrtCnt) = 0, "-", arrRecords3(22, PrtCnt) & "") ' TIM_IMPACT_RST_AVE
            pAry2(LneCnt, 9) = IIf(arrRecords3(22, PrtCnt) = 0, "-", arrRecords3(22, PrtCnt) & "") ' TIM_IMPACT_RST_AVE
            pAry2(LneCnt, 10) = IIf(arrRecords3(22, PrtCnt) = 0, "-", arrRecords3(22, PrtCnt) & "") ' TIM_IMPACT_RST_AVE
        End If
        
        pAry2(LneCnt, 17) = IIf(arrRecords3(23, PrtCnt) = 0, "-", arrRecords3(23, PrtCnt) & "") ' RA_RST_1
        pAry2(LneCnt, 18) = IIf(arrRecords3(24, PrtCnt) = 0, "-", arrRecords3(24, PrtCnt) & "") ' RA_RST_2
        pAry2(LneCnt, 19) = IIf(arrRecords3(25, PrtCnt) = 0, "-", arrRecords3(25, PrtCnt) & "") ' RA_RST_3
        pAry2(LneCnt, 20) = IIf(arrRecords3(26, PrtCnt) = 0, "-", arrRecords3(26, PrtCnt) & "") ' RA_RST_AVE
        pAry2(LneCnt, 21) = IIf(arrRecords3(27, PrtCnt) = 0, "-", arrRecords3(27, PrtCnt) & "") ' GRAIN SIZE
        pAry2(LneCnt, 22) = IIf(arrRecords3(30, PrtCnt) = 0, "-", arrRecords3(30, PrtCnt) & "") ' HARD RSLT
        
        
        
' 2007.4.16 B PRINT END L.Q.Y.
        If STLGRD = "A515Gr60" Or STLGRD = "A515Gr50" Or STLGRD = "A516Gr60" Then
            pAry2(LneCnt, 23) = arrRecords3(28, PrtCnt) & ""             'NON_METAIL
        End If
        
        If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P092" Then
            pAry2(LneCnt, 23) = arrRecords3(28, PrtCnt) & ""             'NON_METAIL
        End If
        
        If arrRecords3(29, PrtCnt) > 0 And STLGRD <> "A515Gr60" Then
           pAry2(LneCnt, 23) = IIf(arrRecords3(29, PrtCnt) = 0, "-", arrRecords3(29, PrtCnt) & "") ' YR_RST
        End If
        
' 2006. 11. 07  INSERT START BY L.Q.Y
        If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P081" Then
           pAry2(LneCnt, 23) = IIf(arrRecords3(29, PrtCnt) = 0, "-", arrRecords3(29, PrtCnt) & "") ' YR_RST
        End If
'-------------------------------------------------------------------ͳ��֧��������----------------------------------------------------

        lSumQNTY = lSumQNTY + arrRecords2(4, PrtCnt)                    'PAGE SUM QUANTITY
        dSumWGT = dSumWGT + arrRecords2(5, PrtCnt)                      'PAGE SUM WEIGHT
        
        lSumQNTY_T = lSumQNTY_T + arrRecords2(4, PrtCnt)                'TOTAL SUM QUANTITY
        dSumWGT_T = dSumWGT_T + arrRecords2(5, PrtCnt)                  'TOTAL SUM WEIGHT
'-------------------------------------------------------------------------------------------------------------------------------------
        If LneCnt = 5 Then
            
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
            
            xlApp.Workbooks.Open (App.Path & "\AQD040C.xls")
            Set xlSheet = xlApp.Worksheets("Sheet1")
'-------------------------------------------------�ʱ���ͷ����ӡ---------------------------------------------------------------------
            Call MillSheetPrint_C_Head(arrRecords1, STLGRD)
'-------------------------------------------------�ʱ��鱸ע��ӡ---------------------------------------------------------------------
            If STLGRD = "S355JR" Or STLGRD = "S355J2" Then
                  xlSheet.Range("U33").Font.Size = 8
                  xlSheet.Range("U33").Value = "Cev=Ceq,Cev=C+Mn/6+(Cr+Mo+V)/5+(Ni+Cu)/15"
            End If

            If Mid(arrRecords1(9, 0), 7, 4) = "P103" And STLGRD = "S355J2G3" Then
                xlSheet.Range("N33").Value = "TMCP ROLLED"
            End If

            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06JTE/P002" Then
                If STLGRD = "A516M485" Then
                    xlSheet.Range("H33").Value = "We confirm the material is manufactured in accordance with the spec. of ASTM A516M485."
                Else
                    xlSheet.Range("H33").Value = "We confirm the material is manufactured in accordance with the spec. of ASTM A283C."
                End If
            End If
            
            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06JTE/P004" Then
                If STLGRD = "S275JR" Then
                    xlSheet.Range("H32").Value = "We confirm the material is manufactured in accordance with the spec. of EN10025 S275JR."
                Else
                    xlSheet.Range("H32").Value = "We confirm the material is manufactured in accordance with the spec. of EN10025 S355J2G3."
                End If
                    xlSheet.Range("H33").Value = "Quality Certificate is in accordance with EN10204 3.1.B ."
            End If
'2008,3,14 sunbin start
             If Mid(Trim(arrRecords1(9, 0)), 1, 13) = "PCLZCPO030035" Then
                    xlSheet.Range("H32").Value = "�������ƣ���֣����Ʒ�͹ܵ����� �а������ƣ���֣��EPC��Ŀ��"

            End If
'2008,3,14 sunbin end

           If Trim(arrRecords1(9, 0)) = "05NGE/P125" Then
                If STLGRD = "A573Gr70" Then
                    Select Case arrRecords2(3, 0)
                        Case "31*2950*11690"
                             xlSheet.Range("N33").Value = "TAW1"
                        Case "24.3*2950*11690"
                             xlSheet.Range("N33").Value = "TAW2"
                        Case "19.2*2950*11690"
                             xlSheet.Range("N33").Value = "TAW3"
                        Case "14*2950*11690"
                             xlSheet.Range("N33").Value = "TAW4"
                        Case "9*2000*10000"
                             xlSheet.Range("N33").Value = "TAW5"
                    End Select
                ElseIf STLGRD = "A36" Then
                    Select Case arrRecords2(3, 0)
                        Case "35*2500*12000"
                             xlSheet.Range("N33").Value = "TAW6"
                        Case "30*3000*12000"
                             xlSheet.Range("N33").Value = "TAW7"
                        Case "20*2500*10000"
                             xlSheet.Range("N33").Value = "TAW8"
                        Case "13*2000*10000"
                             xlSheet.Range("N33").Value = "TAW9"
                        Case "12*2000*11690"
                             xlSheet.Range("N33").Value = "TAW10"
                        Case "11*2950*11690"
                             xlSheet.Range("N33").Value = "TAW11"
                        Case "10*2930*11690"
                             xlSheet.Range("N33").Value = "TAW12"
                        Case "6.5*2000*10000"
                             xlSheet.Range("N33").Value = "TAW13"
                    End Select
                End If
               xlSheet.Range("H33").Value = "AS PER 3.1.B"
            End If

'---------------------------------------------ҳ֧���������ϼ�/��֧���������ϼ�/ҳ���ӡ-----------------------------------------------------------------
            xlSheet.Range("B32").Value = lSumQNTY      'arrRecords1(17, 0) & ""      'SUM_CNT
            xlSheet.Range("C32").Value = dSumWGT       'arrRecords1(18, 0) & ""      'SUM_WGT
            
            If PrtCnt = RowCNT Then
                xlSheet.Range("U32").Value = lSumQNTY_T & " Piece"        'SUM_CNT
                xlSheet.Range("W32").Value = dSumWGT_T & " ton"        'SUM_WGT
            End If
            
            
            xlSheet.Range("L34").Value = Round((RowCNT + 3) / 5, 0) & " - " & Round((PrtCnt + 3) / 5, 0)
            Page_no = Round((PrtCnt + 3) / 5, 0)
'--------------------------------------------------------------------------------------------------------------------------------------------------------
            If STLGRD = "A515Gr60" Or STLGRD = "A515Gr50" Or STLGRD = "A516Gr60" Then
                xlSheet.Range("X22").Value = "Nonm_etal Lard " & ""            'NON_METAIL"
                xlSheet.Range("X24").Value = " " & ""
            End If
            
            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P092" Then
                xlSheet.Range("X22").Value = "Nonm_etal Lard " & ""            'NON_METAIL"
                xlSheet.Range("X24").Value = " " & ""            'NON_METAIL"
            End If
                        
            If arrRecords3(29, PrtCnt) > 0 And STLGRD <> "A515Gr60" Then
               xlSheet.Range("X22").Value = "Y.S./T.S. (%)" & ""
            End If
            
            
'' 2006. 11. 08  INSERT START BY L.Q.Y
            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P081" Then
               xlSheet.Range("X22").Value = "Y.S./T.S. (%)" & ""
            End If
' 2006. 11  .08 INSERT END BY L.Q.Y


            If GetCertSTD(Cert_No) = True Or Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P073" Then
                  xlSheet.SHAPES(1).Visible = True
                  xlSheet.Range("Y32").Value = "0038"
            Else
                  xlSheet.SHAPES(1).Visible = False
                  xlSheet.Range("Y32").Value = ""
            End If
            
            xlSheet.Range("B17:G21").Value = pAry11
            xlSheet.Range("I15:Z15").Value = pAry12
            xlSheet.Range("I16:Z16").Value = pAry13
            xlSheet.Range("I17:Z21").Value = pAry14
            xlSheet.Range("Y22:Z22").Value = pAry15
            xlSheet.Range("Y24:Z24").Value = pAry16
            xlSheet.Range("Y25:Z29").Value = pAry15
            xlSheet.Range("B25:X29").Value = pAry2
            
            Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, iSave_State, sSave_Path)
            If Save_State = 0 Or Save_State = 1 Then

                If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P040" Then
                    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True
                End If

                If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P049" Then
                    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True
                End If

                xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            End If
'
            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit

            LneCnt = 0
            lSumQNTY = 0
            dSumWGT = 0

            
            ReDim pAry11(1 To 5, 1 To 6)                                    '��Ʒ��Ϣ(��Ʒ��\�ƺ�\�ߴ�\֧��\����)
            ReDim pAry12(1 To 1, 1 To 18)                                   '�ɷ�����
            ReDim pAry13(1 To 1, 1 To 18)                                   '�ɷֱ���
            ReDim pAry14(1 To 5, 1 To 18)                                   '�ɷ�ʵ��
            ReDim pAry15(1 To 1, 1 To 2)                                    '���óɷ�����
            ReDim pAry16(1 To 1, 1 To 2)                                    '���óɷֱ���
            ReDim pAry17(1 To 5, 1 To 2)                                    '���óɷ�ʵ��
            
            ReDim pAry2(1 To 5, 1 To 23)                                    '����ʵ��(��һλ��Ʒ��)
            
        End If

    Loop Until PrtCnt = RowCNT
    
'--------------------------------------------------------------------------------------------------------
'-                                            ��ĩҳ��ӡ                                                -
'--------------------------------------------------------------------------------------------------------
    If LneCnt <> 0 Then
    
        
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
            
            xlApp.Workbooks.Open (App.Path & "\AQD040C.xls")
            Set xlSheet = xlApp.Worksheets("Sheet1")
'-------------------------------------------------�ʱ���ͷ����ӡ---------------------------------------------------------------------
            Call MillSheetPrint_C_Head(arrRecords1, STLGRD)
'-------------------------------------------------�ʱ��鱸ע��ӡ---------------------------------------------------------------------
            If STLGRD = "S355JR" Or STLGRD = "S355J2" Then
                  xlSheet.Range("U33").Value = "Ceq = Cev , Cev=C+Mn/6+(Cr+Mo+V)/5+(Ni+Cu)/15"
            End If

            If Mid(arrRecords1(9, 0), 7, 4) = "P103" And STLGRD = "S355J2G3" Then
                xlSheet.Range("N33").Value = "TMCP ROLLED"
            End If

            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06JTE/P002" Then
                If STLGRD = "A516M485" Then
                    xlSheet.Range("H33").Value = "We confirm the material is manufactured in accordance with the spec. of ASTM A516M485."
                Else
                    xlSheet.Range("H33").Value = "We confirm the material is manufactured in accordance with the spec. of ASTM A283C."
                End If
            End If
            
            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06JTE/P004" Then
                If STLGRD = "S275JR" Then
                    xlSheet.Range("H32").Value = "We confirm the material is manufactured in accordance with the spec. of EN10025 S275JR."
                Else
                    xlSheet.Range("H32").Value = "We confirm the material is manufactured in accordance with the spec. of EN10025 S355J2G3."
                End If
                    xlSheet.Range("H33").Value = "Quality Certificate is in accordance with EN10204 3.1.B ."
            End If
            
'2008,4,29 sunbin start
             If Mid(Trim(arrRecords1(9, 0)), 1, 13) = "PCLZCPO030035" Then
                    xlSheet.Range("H32").Value = "�������ƣ���֣����Ʒ�͹ܵ����� �а������ƣ���֣��EPC��Ŀ��"

            End If
'2008,4,29 sunbin end

                       
           If Trim(arrRecords1(9, 0)) = "05NGE/P125" Then
                If STLGRD = "A573Gr70" Then
                    Select Case arrRecords2(3, 0)
                        Case "31*2950*11690"
                             xlSheet.Range("N33").Value = "TAW1"
                        Case "24.3*2950*11690"
                             xlSheet.Range("N33").Value = "TAW2"
                        Case "19.2*2950*11690"
                             xlSheet.Range("N33").Value = "TAW3"
                        Case "14*2950*11690"
                             xlSheet.Range("N33").Value = "TAW4"
                        Case "9*2000*10000"
                             xlSheet.Range("N33").Value = "TAW5"
                    End Select
                ElseIf STLGRD = "A36" Then
                    Select Case arrRecords2(3, 0)
                        Case "35*2500*12000"
                             xlSheet.Range("N33").Value = "TAW6"
                        Case "30*3000*12000"
                             xlSheet.Range("N33").Value = "TAW7"
                        Case "20*2500*10000"
                             xlSheet.Range("N33").Value = "TAW8"
                        Case "13*2000*10000"
                             xlSheet.Range("N33").Value = "TAW9"
                        Case "12*2000*11690"
                             xlSheet.Range("N33").Value = "TAW10"
                        Case "11*2950*11690"
                             xlSheet.Range("N33").Value = "TAW11"
                        Case "10*2930*11690"
                             xlSheet.Range("N33").Value = "TAW12"
                        Case "6.5*2000*10000"
                             xlSheet.Range("N33").Value = "TAW13"
                    End Select
                End If
               xlSheet.Range("H33").Value = "AS PER 3.1.B"
            End If

'---------------------------------------------ҳ֧���������ϼ�/��֧���������ϼ�/ҳ���ӡ-----------------------------------------------------------------
            xlSheet.Range("B32").Value = lSumQNTY      'arrRecords1(17, 0) & ""      'SUM_CNT
            xlSheet.Range("C32").Value = dSumWGT       'arrRecords1(18, 0) & ""      'SUM_WGT
            
            xlSheet.Range("U32").Value = lSumQNTY_T & " Piece"        'SUM_CNT
            xlSheet.Range("W32").Value = dSumWGT_T & " ton"        'SUM_WGT
            
            
            xlSheet.Range("L34").Value = Round((RowCNT + 3) / 5, 0) & " - " & Round((PrtCnt + 3) / 5, 0)
            Page_no = Round((PrtCnt + 3) / 5, 0)
'--------------------------------------------------------------------------------------------------------------------------------------------------------
            If STLGRD = "A515Gr60" Or STLGRD = "A515Gr50" Or STLGRD = "A516Gr60" Then
                xlSheet.Range("X22").Value = "Nonm_etal Lard " & ""            'NON_METAIL"
                xlSheet.Range("X24").Value = " " & ""
            End If
            
            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P092" Then
                xlSheet.Range("X22").Value = "Nonm_etal Lard " & ""            'NON_METAIL"
                xlSheet.Range("X24").Value = " " & ""            'NON_METAIL"
            End If
                        
            If arrRecords3(29, PrtCnt) > 0 And STLGRD <> "A515Gr60" Then
               xlSheet.Range("X22").Value = "Y.S./T.S. (%)" & ""
            End If
            
            
'' 2006. 11. 08  INSERT START BY L.Q.Y
            If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P081" Then
               xlSheet.Range("X22").Value = "Y.S./T.S. (%)" & ""
            End If
' 2006. 11  .08 INSERT END BY L.Q.Y



            If GetCertSTD(Cert_No) = True Or Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P073" Then
                  xlSheet.SHAPES(1).Visible = True
                  xlSheet.Range("Y32").Value = "0038"
            Else
                  xlSheet.SHAPES(1).Visible = False
                  xlSheet.Range("Y32").Value = ""
            End If
            
            xlSheet.Range("B17:G21").Value = pAry11
            xlSheet.Range("I15:Z15").Value = pAry12
            xlSheet.Range("I16:Z16").Value = pAry13
            xlSheet.Range("I17:Z21").Value = pAry14
            xlSheet.Range("Y22:Z22").Value = pAry15
            xlSheet.Range("Y24:Z24").Value = pAry16
            xlSheet.Range("Y25:Z29").Value = pAry15
            xlSheet.Range("B25:X29").Value = pAry2
            
            Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, iSave_State, sSave_Path)
            If Save_State = 0 Or Save_State = 1 Then

                If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P040" Then
                    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True
                End If

                If Mid(Trim(arrRecords1(9, 0)), 1, 10) = "06NGE/P049" Then
                    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True
                End If

                xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            End If
'
            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit

    End If
    
    Set xlApp = Nothing
    
    Exit Function
    
End Function


'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_C_Head
'   2.Name         : Conventionality certificate print(Head table)
'   3.Input  Value : arrRecords1,sStlgrd
'   4.Return Value :
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Sub MillSheetPrint_C_Head(ByVal arrRecords1 As Variant, ByVal sSTLGRD As String)
    Dim s_MyPono        As String
    Dim sDate           As String
    Dim sCert_No        As String
    Dim sPROD_CD        As String
    Dim sCONDITION_CN   As String
    Dim sCONDITION_EN   As String
    Dim sREMARK         As String
    Dim sSTDSPEC        As String
    Dim sUSTSTDNAME     As String
    Dim sLicence_NO     As String
    Dim sISP_SHP_NO     As String
        
        
        
    sDate = GetPrintDate()
    sCert_No = arrRecords1(0, 0)
    sPROD_CD = Mid(sCert_No, 1, 2)
    sISP_SHP_NO = arrRecords1(10, 0)
    sCONDITION_CN = GetConditionOfDelivery_C(sISP_SHP_NO)
    sCONDITION_EN = GetConditionOfDelivery_E(sISP_SHP_NO)
    s_MyPono = Trim(CERT_PONO_SET(sISP_SHP_NO))
    
    If s_MyPono = "NO" Then
        s_MyPono = GetPoNoLot(sISP_SHP_NO)
        If s_MyPono = "N" Then
            If Trim(arrRecords1(0, 0) & "") = "PP200606230022" Or _
               Trim(arrRecords1(0, 0) & "") = "PP200606230023" Then
                s_MyPono = "06NGE/P043LOT5"
            Else
                If IsNull(arrRecords1(9, 0)) Then
                    s_MyPono = ""
                Else
                    s_MyPono = arrRecords1(9, 0)
                End If
            End If
        Else
            s_MyPono = s_MyPono        'PONO
        End If
    Else
        s_MyPono = s_MyPono        'PONO
    End If
    
    If Trim(arrRecords1(2, 0)) = "2006BJ17-2006" Then
        sSTDSPEC = "1E0170����Э��"
    ElseIf s_MyPono = "YJ0702018" And Trim(Mid(arrRecords1(2, 0), 1, 4)) = "ASTM" Then
        sSTDSPEC = "ASTM A709-2000"
    ElseIf s_MyPono = "07NGE/P016" Or s_MyPono = "07NGE/P013" Or _
        s_MyPono = "07JTE/P004-1" Or s_MyPono = "07NGE/P011" Then
        sSTDSPEC = "EN 10025-2:2004"
    ElseIf IsNull(arrRecords1(2, 0)) Then
        sSTDSPEC = ""
    Else
        sSTDSPEC = arrRecords1(2, 0)
    End If
    
    sREMARK = Old_Remark(sSTDSPEC, s_MyPono)
    If sREMARK = "N" Then
        sREMARK = GetRemark(sISP_SHP_NO)
        If sREMARK = "N" Then
            If Mid(sSTDSPEC, 1, 2) = "EN" Then
                sREMARK = "According to EN 10204 3.1.B"
            Else
                sREMARK = ""
            End If
        End If
    Else
        sREMARK = sREMARK
    End If
    
    If arrRecords1(17, 0) = "EN10160 E23" Then
        sUSTSTDNAME = "EN10160 S2E3"
    ElseIf IsNull(arrRecords1(17, 0)) Then
        sUSTSTDNAME = ""
    Else
        sUSTSTDNAME = arrRecords1(17, 0)
    End If
    
    If (sCONDITION_CN = "N" Or Len(Trim(sCONDITION_CN)) = 0) Or _
       (sCONDITION_EN = "N" Or Len(Trim(sCONDITION_EN)) = 0) Then
        sCONDITION_EN = arrRecords1(5, 0)
        If Left(arrRecords1(5, 0), 1) = "H" Then
            sCONDITION_CN = "����"
        ElseIf Left(arrRecords1(5, 0), 1) = "C" Then
            sCONDITION_CN = "����"
        ElseIf Left(arrRecords1(5, 0), 1) = "N" Then
            sCONDITION_CN = "��������"
        ElseIf Left(arrRecords1(5, 0), 1) = "T" Then
            sCONDITION_CN = "�ػ�"
        ElseIf Left(arrRecords1(5, 0), 1) = "Q" Then
            sCONDITION_CN = "���"
        End If
    End If
' 2008,0501  sunbin  start
'    If sSTLGRD = "20g" Or sSTLGRD = "16Mng" Or sSTLGRD = "19Mng" Or _
'        sSTLGRD = "22Mng" Or sSTLGRD = "15CrMog" Or sSTLGRD = "20R" Or _
'        sSTLGRD = "16MnR" Or sSTLGRD = "15MnVR" Then
'        sLicence_NO = "�����豸��������֤��" & arrRecords1(3, 0)
'    ElseIf IsNull(arrRecords1(3, 0)) Then
'        sLicence_NO = ""
'    Else
'        sLicence_NO = arrRecords1(3, 0)
'    End If
'2008.0501   sunbin end
    If IsNull(arrRecords1(3, 0)) Then
        sLicence_NO = ""
    Else
        sLicence_NO = arrRecords1(3, 0)
    End If
   
'---------------- print ---------------------------------------------------------------------------------
    xlSheet.Range("C2").Value = sCert_No & ""        'CERT_NO
    xlSheet.Range("C4").Value = arrRecords1(1, 0) & ""        'PROD_NAME
    xlSheet.Range("C5").Value = arrRecords1(21, 0) & ""        'PROD_NAME_ENG
    xlSheet.Range("C6").Value = sSTDSPEC & ""        'STDSPEC_NAME
    xlSheet.Range("C8").Value = sLicence_NO & ""        'PROD_SPEC_NO
    
    xlSheet.Range("C10").Value = arrRecords1(4, 0) & ""       'CUST_NAME
    
    xlSheet.Range("B13").Value = "Condition of Supply�� " & sCONDITION_EN & ""       'COND_SUPPLY OF ENGLISH
    xlSheet.Range("C12").Value = sCONDITION_CN & ""                                  'COND_SUPPLY OF CHINESE
    xlSheet.Range("N9").Value = sUSTSTDNAME & ""       'UST_NAME
    xlSheet.Range("N11").Value = Trim(arrRecords1(7, 0)) & "" 'IMPACT_SMP_SIZE
    xlSheet.Range("V4").Value = s_MyPono        'PONO
    xlSheet.Range("V6").Value = arrRecords1(10, 0) & ""       'TRNS_NO
    xlSheet.Range("V8").Value = arrRecords1(11, 0) & ""       'TRAIN_LINE_NAME
    xlSheet.Range("V10").Value = ""                           'DEST_DETAIL
    xlSheet.Range("V12").Value = sDate                        'CERT_RPT_DATE
    If InStr(1, Trim(s_MyPono), "������") > 0 Or InStr(1, Trim(s_MyPono), "YJ0801008") > 0 Then
        xlSheet.Range("J31").Value = "�ʼ츺����" + Chr$(13) + Chr$(10) + "Chief  Inspector"
    ElseIf InStr(1, Trim(s_MyPono), "������") > 0 Or InStr(1, Trim(s_MyPono), "YJ0802002") > 0 Then
        xlSheet.Range("J31").Value = "�ʼ츺����" + Chr$(13) + Chr$(10) + "Chief  Inspector"
    ElseIf InStr(1, Trim(s_MyPono), "������") > 0 Or InStr(1, Trim(s_MyPono), "YJ0803008") > 0 Then
        xlSheet.Range("J31").Value = "�ʼ츺����" + Chr$(13) + Chr$(10) + "Chief  Inspector"
    Else
        xlSheet.Range("M31").Value = sUserName      'TEST_EMP
    End If
'    xlSheet.Range("M31").Value = sUsername      'TEST_EMP
    xlSheet.Range("R31").Value = arrRecords1(16, 0) & ""      'SHP_EMP
    xlSheet.Range("H33").Value = sREMARK
'---------------------  Э�� -----------------------------------------
    If Trim(arrRecords1(20, 0) & "") = "3" Then
        xlSheet.Range("B2").Value = "Э����ţ�"
        xlSheet.Range("B3").Value = "Concert No."
        xlSheet.Range("C4").Value = arrRecords1(1, 0) & "(Э��Ʒ)"        'PROD_NAME
        xlSheet.Range("I5").Value = "Э �� ֤ �� ��"
        xlSheet.Range("H7").Value = "Agreement Certificate"
        xlSheet.Range("C6").Value = "�û�Э��"
        xlSheet.Range("D30").Value = "Э���Ʒ"
        xlSheet.Range("D31").Value = "����˵����"
    End If
    
    If UCase(sPROD_CD) = "HC" Then
        xlSheet.Range("B14").Value = "��Ʒ��" + Chr$(10) + "Coil ��"
        xlSheet.Range("B22").Value = "��Ʒ��" + Chr$(10) + "Coil ��"
    Else
        xlSheet.Range("B14").Value = "��Ʒ��" + Chr$(10) + "Plate ��"
        xlSheet.Range("B22").Value = "��Ʒ��" + Chr$(10) + "Plate ��"
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - funGetQuery_S
'   2.Name         : Ship certificate
'   3.Input  Value : sCertNo , iSave_State ,sSave_Path
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function funGetQuery_S(ByVal sCertNo As String, ByVal iSave_State As Integer, ByVal sSave_Path As String)
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
    Dim AdoRs As adodb.Recordset
    Dim sPROD_CD As String
    Dim sTable_PROD As String
    Dim sFieldName_NO As String
    
    sPROD_CD = Mid(sCertNo, 1, 2)
    If UCase(sPROD_CD) = "HC" Then
        sTable_PROD = "GP_COIL"
        sFieldName_NO = "COIL_NO"
    Else
        sTable_PROD = "GP_PLATE"
        sFieldName_NO = "PLATE_NO"
    End If
    
    Set AdoRs = New adodb.Recordset
'    ���ӿͻ����� ��ѧ�� 2011 0615
'    ���Ӷ�����ע�����ݶ�����ע�Ƿ��С�ִ���¹淶�������Ƿ��ӡ�ɴ���ע�ͣ� ����   2011 1009
'    ����Ӳ������Ҫ��,���� ��ѧ�� 20120221
     sQuery = "SELECT CERT_NO  , PROD_SPEC_NO , STDSPEC_NAME , COND_SUPPLY ,Gf_Cert_Spec(STD_ORGAN, A.STDSPEC) , IMPACT_SMP_SIZE "
    sQuery = sQuery + ", SHIP_CMPY_NO , CERT_RPT_DATE , Gf_Cert_Org(STD_ORGAN) , GF_EMPNAMEFIND(TEST_EMP) AS TEST_EMP "
    sQuery = sQuery + ", AQD0060C.F_SUM_CNT(CERT_NO) AS SUM_CNT, AQD0060C.F_SUM_WGT(CERT_NO) AS SUM_WGT,DECODE(STD_ORGAN,'CCS','',CONTROL_NO),GF_AQD0010_TXT1(STD_ORGAN,1)"  '14
    sQuery = sQuery + ",GF_AQD0010_TXT1(STD_ORGAN,2),GF_AQD0010_TXT1(STD_ORGAN,3),GF_AQD0010_TXT1(STD_ORGAN,4),GF_AQD0010_TXT1(STD_ORGAN,5)"
    sQuery = sQuery + ",GF_AQD0010_TXT1(STD_ORGAN,6),GF_AQD0010_TXT1(STD_ORGAN,7),GF_AQD0010_TXT1(STD_ORGAN,8),GF_AQD0010_TXT1(STD_ORGAN,9)"
    sQuery = sQuery + ",GF_AQD0010_TXT1(STD_ORGAN,10) , QLTY_REC_NO,PROD_SIZE, GF_UST_STD(A.UST_FL),ORD_CUST_NM,B.COLOR_STROKE ,GF_AQD0010_TXT2(A.ORD_NO,A.ORD_ITEM,3),GF_AQD0010_TXT2(A.ORD_NO,A.ORD_ITEM,2),GF_AQD0010_TXT2(A.ORD_NO,A.ORD_ITEM,1) "
    sQuery = sQuery + ",B.VESSEL_NO,B.JIT_FLAG,A.ORD_NO,A.ORD_ITEM"
    sQuery = sQuery + " FROM QP_CERT_HEAD A, BP_ORDER_ITEM B WHERE A.CERT_NO  = '" & sCertNo & "' "
    sQuery = sQuery + " AND A.ORD_NO = B.ORD_NO AND A.ORD_ITEM = B.ORD_ITEM "
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_S = "Err DataBase"
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close
'GF_MARKING_SMP(CERT_NO,SMP_NO)
    sQuery = "SELECT A.CERT_NO ,A.SMP_NO , A.PROD_SIZE , A.PRDT_QNTY , A.PRDT_WGT"
    sQuery = sQuery + ", DECODE(A.C_RST,NULL,0,A.C_RST) , DECODE(A.SI_RST,NULL,0,A.SI_RST) , DECODE(A.MN_RST,NULL,0,A.MN_RST) , DECODE(A.P_RST,NULL,0,A.P_RST) "
    sQuery = sQuery + ", DECODE(A.S_RST,NULL,0,A.S_RST),DECODE(A.NB_RST,NULL,0,A.NB_RST) , DECODE(A.AL_RST,NULL,0,A.AL_RST) , DECODE(A.MO_RST,NULL,0,A.MO_RST) "
    sQuery = sQuery + ", DECODE(A.CU_RST,NULL,0,A.CU_RST) , DECODE(A.NI_RST,NULL,0,A.NI_RST), DECODE(A.CR_RST,NULL,0,A.CR_RST),DECODE(A.TI_RST,NULL,0,A.TI_RST)"
    sQuery = sQuery + ", DECODE(A.CEQ_RST,NULL,0,A.CEQ_RST),DECODE(A.YP_RST,NULL,0,A.YP_RST) , DECODE(A.TS_RST,NULL,0,A.TS_RST), DECODE(A.EL_RST,NULL,0,A.EL_RST)"
    sQuery = sQuery + ", DECODE(A.IMPACT_TMP,NULL,0,A.IMPACT_TMP), DECODE(A.IMPACT_RST1,NULL,0,A.IMPACT_RST1), DECODE(A.IMPACT_RST2,NULL,0,A.IMPACT_RST2)"
    sQuery = sQuery + ", DECODE(A.IMPACT_RST3,NULL,0,A.IMPACT_RST3), DECODE(A.IMPACT_RST4,NULL,0,A.IMPACT_RST4), DECODE(A.IMPACT_RST5,NULL,0,A.IMPACT_RST5)"
    sQuery = sQuery + ", DECODE(A.IMPACT_RST6,NULL,0,A.IMPACT_RST6),DECODE(A.IMPACT_RST_AVE,NULL,0,A.IMPACT_RST_AVE),A.PROD_NO_TEXT"
    sQuery = sQuery + ", GF_AQD0060C_FIND('V',A.PROD_NO,A.CERT_NO),GF_AQD0060C_FIND('N',A.PROD_NO,A.CERT_NO)"
    sQuery = sQuery + ", DECODE(B.A_IMPACT_TMP,NULL,0,B.A_IMPACT_TMP), DECODE(B.A_IMPACT_RST1,NULL,0,B.A_IMPACT_RST1), DECODE(B.A_IMPACT_RST2,NULL,0,B.A_IMPACT_RST2)"
    sQuery = sQuery + ", DECODE(B.A_IMPACT_RST3,NULL,0,B.A_IMPACT_RST3), DECODE(B.A_IMPACT_RST4,NULL,0,B.A_IMPACT_RST4), DECODE(B.A_IMPACT_RST5,NULL,0,B.A_IMPACT_RST5)"
    sQuery = sQuery + ", DECODE(B.A_IMPACT_RST6,NULL,0,B.A_IMPACT_RST6),DECODE(B.A_IMPACT_RST_AVE,NULL,0,B.A_IMPACT_RST_AVE),DECODE(B.IMPACT_DIR,NULL,'',B.IMPACT_DIR),DECODE(B.A_IMPACT_DIR,NULL,'',B.A_IMPACT_DIR) "
    sQuery = sQuery + ", B.ZRA_RST1,B.ZRA_RST2,B.ZRA_RST3,B.ZRA_RST_AVE"
    sQuery = sQuery + ", B.IMPACT_DIR,B.A_IMPACT_DIR,B.HARD_RST,B.BEND_RST,DECODE(A.ZR_RST,NULL,0,A.ZR_RST), DECODE(A.B_RST,NULL,0,A.B_RST)"
    sQuery = sQuery + " FROM QP_CERT_DETAIL A , QP_TEST_RSLT B WHERE A.SMP_NO=B.SMP_NO AND A.CERT_NO  = '" & sCertNo & "'"
           
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_S = "Err DataBase"
        Exit Function
    End If
    arrRecords2 = AdoRs.GetRows
    AdoRs.Close
    
    funGetQuery_S = MillSheetPrint_S(iSave_State, sSave_Path, arrRecords1, arrRecords2)
    Set AdoRs = Nothing
    
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_S
'   2.Name         : Ship certificate print(Detail table)
'   3.Input  Value : iSave_State ,sSave_Path ,arrRecords1 ,arrRecords2
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function MillSheetPrint_S(ByVal iSave_State As Integer, ByVal sSave_Path As String, ByVal arrRecords1 As Variant, ByVal arrRecords2 As Variant) As String
    Dim RowCNT      As Long
    Dim PrtCnt      As Long
    Dim LneCnt      As Long
    Dim pAry()      As String
    Dim pAry1()     As String
    Dim iDx         As Integer
    Dim sRow        As String
    Dim lSumQNTY    As Long
    Dim dSumWGT     As Double
    Dim lSumQNTY_T    As Long
    Dim dSumWGT_T     As Double
    Dim Save_State  As Integer
    Dim Save_Path As String
    Dim Page_no As String
    Dim Cert_No         As String
    Dim LneCnt_D As Long
    Dim sADD_MATR      As String
        
    Dim dAry()      As Double  '   Ϊ�������� ��ֹС��1t ���� .XXX
       
    Save_State = iSave_State
    Save_Path = sSave_Path
    Cert_No = arrRecords2(0, 0)
    
    If IsEmpty(arrRecords1) Or IsEmpty(arrRecords2) Then
        MillSheetPrint_S = "Err data"
        Exit Function
    End If
    
    RowCNT = UBound(arrRecords2, 2)
    
    PrtCnt = -1
    LneCnt = -1
    lSumQNTY = 0
    dSumWGT = 0
    lSumQNTY_T = 0
    dSumWGT_T = 0
    LneCnt_D = 0
    
'    If arrRecords1(29, 0) = "A" Then
'        sADD_MATR = "B"                        '  �����ӡ
'    End If
    
    If arrRecords1(28, 0) = "A" Then
        sADD_MATR = "A"                       '   Ӳ��
    Else
        sADD_MATR = ""
    End If
        
    ReDim pAry(1 To 6, 1 To 38)
    ReDim dAry(1 To 6, 1 To 38)
    ReDim pAry1(1 To 6, 1 To 2)
    
    Do

        LneCnt = LneCnt + 2
        PrtCnt = PrtCnt + 1
   
        pAry(LneCnt, 1) = arrRecords2(1, PrtCnt) & ""                  ' PROD_NAME
        pAry(LneCnt, 2) = Mid(arrRecords2(1, PrtCnt), 1, 8) & ""       ' ¯��
        pAry(LneCnt, 3) = arrRecords2(3, PrtCnt) & ""                  ' PRDT_QNTY
'        ��ѧ��  ���
        pAry(LneCnt, 4) = arrRecords2(2, PrtCnt) & ""                      ' ���
'        ��ֹ����С��1t ����.XXX ���Ұ�ģ���Ϊ����
        pAry(LneCnt, 5) = IIf(arrRecords2(4, PrtCnt) < 1, "0" & arrRecords2(4, PrtCnt), _
        arrRecords2(4, PrtCnt))                                        ' PRDT_WGT
        
        pAry(LneCnt, 6) = IIf(Val(arrRecords2(5, PrtCnt) & "") = 0, _
        "-", IIf(Mid(Val(arrRecords2(5, PrtCnt) & "") * 100, 1, 1) = ".", _
        "0" & Val(arrRecords2(5, PrtCnt) & "") * 100, Val(arrRecords2(5, PrtCnt) & "") * 100))               ' C_RST
        
        pAry(LneCnt, 7) = IIf(Val(arrRecords2(6, PrtCnt) & "") = 0, _
        "-", IIf(Mid(Val(arrRecords2(6, PrtCnt) & "") * 100, 1, 1) = ".", _
        "0" & Val(arrRecords2(6, PrtCnt) & "") * 100, Val(arrRecords2(6, PrtCnt) & "") * 100))                    ' SI_RST
        
        pAry(LneCnt, 8) = IIf(Val(arrRecords2(7, PrtCnt) & "") = 0, _
        "-", IIf(Mid(Val(arrRecords2(7, PrtCnt) & "") * 100, 1, 1) = ".", _
        "0" & Val(arrRecords2(7, PrtCnt) & "") * 100, Val(arrRecords2(7, PrtCnt) & "") * 100))                    ' MN_RST
        
        pAry(LneCnt, 9) = IIf(Val(arrRecords2(8, PrtCnt) & "") = 0, _
        "-", IIf(Mid(Val(arrRecords2(8, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(8, PrtCnt) & "") * 1000, Val(arrRecords2(8, PrtCnt) & "") * 1000))                 ' P_RST
        
        pAry(LneCnt, 10) = IIf(Val(arrRecords2(9, PrtCnt) & "") = 0, _
        "-", IIf(Mid(Val(arrRecords2(9, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(9, PrtCnt) & "") * 1000, Val(arrRecords2(9, PrtCnt) & "") * 1000))                 ' S_RST
        
        pAry(LneCnt, 11) = IIf(Val(arrRecords2(10, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(10, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(10, PrtCnt) & "") * 1000, Val(arrRecords2(10, PrtCnt) & "") * 1000))                ' NB_RST
        
        pAry(LneCnt, 12) = IIf(Val(arrRecords2(11, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(11, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(11, PrtCnt) & "") * 1000, Val(arrRecords2(11, PrtCnt) & "") * 1000))                ' AL_RST
        
        pAry(LneCnt, 13) = IIf(Val(arrRecords2(12, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(12, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(12, PrtCnt) & "") * 1000, Val(arrRecords2(12, PrtCnt) & "") * 1000))                ' MO_RST
        
        pAry(LneCnt, 14) = IIf(Val(arrRecords2(13, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(13, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(13, PrtCnt) & "") * 1000, Val(arrRecords2(13, PrtCnt) & "") * 1000))                 ' CU_RST
        
        pAry(LneCnt, 15) = IIf(Val(arrRecords2(14, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(14, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(14, PrtCnt) & "") * 1000, Val(arrRecords2(14, PrtCnt) & "") * 1000))                 ' NI_RST
        
        pAry(LneCnt, 16) = IIf(Val(arrRecords2(15, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(15, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(15, PrtCnt) & "") * 1000, Val(arrRecords2(15, PrtCnt) & "") * 1000))                 ' CR_RST
        
        pAry(LneCnt, 17) = IIf(Val(arrRecords2(16, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(16, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(16, PrtCnt) & "") * 1000, Val(arrRecords2(16, PrtCnt) & "") * 1000))                 ' TI_RST
        
        pAry(LneCnt, 18) = IIf(Val(arrRecords2(30, PrtCnt) & "") = 0, _
        "<1", IIf(Mid(Val(arrRecords2(30, PrtCnt) & "") * 1000, 1, 1) = ".", _
        "0" & Val(arrRecords2(30, PrtCnt) & "") * 1000, Val(arrRecords2(30, PrtCnt) & "") * 1000))                ' V_RST
        
        pAry(LneCnt, 19) = IIf(Val(arrRecords2(31, PrtCnt) & "") = 0, _
        "-", IIf(Mid(Val(arrRecords2(31, PrtCnt) & ""), 1, 1) = ".", _
        "0" & Val(arrRecords2(31, PrtCnt) & ""), Val(arrRecords2(31, PrtCnt) & "")))                   ' N_RST
        
        
        pAry(LneCnt, 20) = IIf(Val(arrRecords2(17, PrtCnt) & "") = 0, _
        "-", IIf(Mid(Val(arrRecords2(17, PrtCnt) & "") * 100, 1, 1) = ".", _
        "0" & Val(arrRecords2(17, PrtCnt) & "") * 100, Val(arrRecords2(17, PrtCnt) & "") * 100))                 ' CEQ_RST
        
        If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
        
           If arrRecords1(3, 0) = "QT" Then
           
               pAry(LneCnt, 21) = IIf(Val(arrRecords2(50, PrtCnt) & "") = 0, _
                "-", IIf(Mid(Val(arrRecords2(50, PrtCnt) & "") * 100, 1, 1) = ".", _
                "0" & Val(arrRecords2(50, PrtCnt) & "") * 100, Val(arrRecords2(50, PrtCnt) & "") * 100))                 ' Zr
                
                pAry(LneCnt, 22) = IIf(Val(arrRecords2(51, PrtCnt) & "") = 0, _
                "-", IIf(Mid(Val(arrRecords2(51, PrtCnt) & "") * 100, 1, 1) = ".", _
                "0" & Val(arrRecords2(51, PrtCnt) & "") * 100, Val(arrRecords2(51, PrtCnt) & "") * 100))                 ' B
           Else
               pAry(LneCnt, 21) = "-"
               pAry(LneCnt, 22) = "-"
           End If
           
            pAry(LneCnt, 23) = IIf(arrRecords2(18, PrtCnt) = 0, "-", arrRecords2(18, PrtCnt) & "")  ' YP_RST
            pAry(LneCnt, 24) = IIf(arrRecords2(19, PrtCnt) = 0, "-", arrRecords2(19, PrtCnt) & "")  ' TS_RST
            pAry(LneCnt, 25) = IIf(arrRecords2(20, PrtCnt) = 0, "-", arrRecords2(20, PrtCnt) & "")  ' EL_RST
            pAry(LneCnt, 26) = IIf(arrRecords2(22, PrtCnt) = 0, "-", arrRecords2(21, PrtCnt) & "")  ' IMPACT_TMP
            pAry(LneCnt, 27) = IIf(arrRecords2(22, PrtCnt) = 0, "-", arrRecords2(22, PrtCnt) & "")  ' IMPACT_RST_1
            pAry(LneCnt, 28) = IIf(arrRecords2(23, PrtCnt) = 0, "-", arrRecords2(23, PrtCnt) & "")  ' IMPACT_RST_2
            pAry(LneCnt, 29) = IIf(arrRecords2(24, PrtCnt) = 0, "-", arrRecords2(24, PrtCnt) & "")  ' IMPACT_RST_3
            pAry(LneCnt, 30) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
            'pAry(LneCnt, 29) = IIf(arrRecords2(22, PrtCnt) = 0, "-", "L")
            pAry(LneCnt, 31) = IIf(arrRecords2(22, PrtCnt) = 0, "-", IIf(arrRecords2(46, PrtCnt) = "V", "L", "T") & "") ' IMPACT_DIR
                  
    
            pAry(LneCnt + 1, 26) = IIf(arrRecords2(33, PrtCnt) = 0, "-", arrRecords2(32, PrtCnt) & "") ' A_IMPACT_TMP
            pAry(LneCnt + 1, 27) = IIf(arrRecords2(33, PrtCnt) = 0, "-", arrRecords2(33, PrtCnt) & "") ' A_IMPACT_RST_1
            pAry(LneCnt + 1, 28) = IIf(arrRecords2(34, PrtCnt) = 0, "-", arrRecords2(34, PrtCnt) & "") ' A_IMPACT_RST_2
            pAry(LneCnt + 1, 29) = IIf(arrRecords2(35, PrtCnt) = 0, "-", arrRecords2(35, PrtCnt) & "") ' A_IMPACT_RST_3
            pAry(LneCnt + 1, 30) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
            'pAry(LneCnt + 1, 29) = IIf(arrRecords2(33, PrtCnt) = 0, "-", "T")
            pAry(LneCnt + 1, 31) = IIf(arrRecords2(33, PrtCnt) = 0, "-", IIf(arrRecords2(47, PrtCnt) = "V", "L", "T") & "") ' A_IMPACT_DIR
            
            If arrRecords2(25, PrtCnt) > 0 Or arrRecords2(26, PrtCnt) > 0 Or arrRecords2(27, PrtCnt) > 0 Then
                pAry(LneCnt, 27) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
                pAry(LneCnt, 28) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
                pAry(LneCnt, 29) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
            End If
            
             If arrRecords2(36, PrtCnt) > 0 Or arrRecords2(37, PrtCnt) > 0 Or arrRecords2(38, PrtCnt) > 0 Then
                pAry(LneCnt + 1, 27) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
                pAry(LneCnt + 1, 28) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
                pAry(LneCnt + 1, 29) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
            End If
            
            '20100611 sun bin atart
            pAry(LneCnt, 32) = IIf(arrRecords2(42, PrtCnt) = 0, "-", arrRecords2(42, PrtCnt) & "")  ' RA_RST1
            pAry(LneCnt, 33) = IIf(arrRecords2(43, PrtCnt) = 0, "-", arrRecords2(43, PrtCnt) & "")  ' RA_RST2
            pAry(LneCnt, 34) = IIf(arrRecords2(44, PrtCnt) = 0, "-", arrRecords2(44, PrtCnt) & "")  ' RA_RST3
            pAry(LneCnt, 35) = IIf(arrRecords2(45, PrtCnt) = 0, "-", arrRecords2(45, PrtCnt) & "")  ' RA_RST_AVE
    '        pAry(LneCnt, 33) = IIf(IsNull(arrRecords1(25, 0)) = True, "-", "OK" & "") '  UT
            pAry(LneCnt, 36) = IIf(arrRecords1(25, 0) = " ", "-", "OK" & "") '  UT

            
    '        �����ʾ׷�ӵ�����Ҫ��  ��ѧ�� 20120221
            
                
            If sADD_MATR = "A" Then
                        pAry(LneCnt, 37) = IIf(arrRecords2(48, PrtCnt) = 0, "-", arrRecords2(48, PrtCnt) & "")    '�����ʱ���׷�Ӵ�ӡӲ��
    '        ElseIf sADD_MATR = "B" Then
    '                    pAry(LneCnt, 35) = IIf(arrRecords2(49, PrtCnt) = 0, "-", arrRecords2(49, PrtCnt) & "")    '  �����ӡ
            Else
                        pAry(LneCnt, 37) = "-"                          ' �հ�-
            End If
        
    '        pAry(LneCnt, 35) = IIf(arrRecords1(28, 0) = "A", arrRecords2(48, PrtCnt) & "", "-")     ' HARD_RST
        Else
            pAry(LneCnt, 21) = IIf(arrRecords2(18, PrtCnt) = 0, "-", arrRecords2(18, PrtCnt) & "")  ' YP_RST
            pAry(LneCnt, 22) = IIf(arrRecords2(19, PrtCnt) = 0, "-", arrRecords2(19, PrtCnt) & "")  ' TS_RST
            pAry(LneCnt, 23) = IIf(arrRecords2(20, PrtCnt) = 0, "-", arrRecords2(20, PrtCnt) & "")  ' EL_RST
            pAry(LneCnt, 24) = IIf(arrRecords2(22, PrtCnt) = 0, "-", arrRecords2(21, PrtCnt) & "")  ' IMPACT_TMP
            pAry(LneCnt, 25) = IIf(arrRecords2(22, PrtCnt) = 0, "-", arrRecords2(22, PrtCnt) & "")  ' IMPACT_RST_1
            pAry(LneCnt, 26) = IIf(arrRecords2(23, PrtCnt) = 0, "-", arrRecords2(23, PrtCnt) & "")  ' IMPACT_RST_2
            pAry(LneCnt, 27) = IIf(arrRecords2(24, PrtCnt) = 0, "-", arrRecords2(24, PrtCnt) & "")  ' IMPACT_RST_3
            pAry(LneCnt, 28) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
            'pAry(LneCnt, 29) = IIf(arrRecords2(22, PrtCnt) = 0, "-", "L")
            pAry(LneCnt, 29) = IIf(arrRecords2(22, PrtCnt) = 0, "-", IIf(arrRecords2(46, PrtCnt) = "V", "L", "T") & "") ' IMPACT_DIR
                  
    
            pAry(LneCnt + 1, 24) = IIf(arrRecords2(33, PrtCnt) = 0, "-", arrRecords2(32, PrtCnt) & "") ' A_IMPACT_TMP
            pAry(LneCnt + 1, 25) = IIf(arrRecords2(33, PrtCnt) = 0, "-", arrRecords2(33, PrtCnt) & "") ' A_IMPACT_RST_1
            pAry(LneCnt + 1, 26) = IIf(arrRecords2(34, PrtCnt) = 0, "-", arrRecords2(34, PrtCnt) & "") ' A_IMPACT_RST_2
            pAry(LneCnt + 1, 27) = IIf(arrRecords2(35, PrtCnt) = 0, "-", arrRecords2(35, PrtCnt) & "") ' A_IMPACT_RST_3
            pAry(LneCnt + 1, 28) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
            'pAry(LneCnt + 1, 29) = IIf(arrRecords2(33, PrtCnt) = 0, "-", "T")
            pAry(LneCnt + 1, 29) = IIf(arrRecords2(33, PrtCnt) = 0, "-", IIf(arrRecords2(47, PrtCnt) = "V", "L", "T") & "") ' A_IMPACT_DIR
            
            If arrRecords2(25, PrtCnt) > 0 Or arrRecords2(26, PrtCnt) > 0 Or arrRecords2(27, PrtCnt) > 0 Then
                pAry(LneCnt, 25) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
                pAry(LneCnt, 26) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
                pAry(LneCnt, 27) = IIf(arrRecords2(28, PrtCnt) = 0, "-", arrRecords2(28, PrtCnt) & "")  ' IMPACT_RST_AVE
            End If
            
             If arrRecords2(36, PrtCnt) > 0 Or arrRecords2(37, PrtCnt) > 0 Or arrRecords2(38, PrtCnt) > 0 Then
                pAry(LneCnt + 1, 25) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
                pAry(LneCnt + 1, 26) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
                pAry(LneCnt + 1, 27) = IIf(arrRecords2(39, PrtCnt) = 0, "-", arrRecords2(39, PrtCnt) & "") ' A_IMPACT_RST_AVE
            End If
            
    '20100611 sun bin atart
            pAry(LneCnt, 30) = IIf(arrRecords2(42, PrtCnt) = 0, "-", arrRecords2(42, PrtCnt) & "")  ' RA_RST1
            pAry(LneCnt, 31) = IIf(arrRecords2(43, PrtCnt) = 0, "-", arrRecords2(43, PrtCnt) & "")  ' RA_RST2
            pAry(LneCnt, 32) = IIf(arrRecords2(44, PrtCnt) = 0, "-", arrRecords2(44, PrtCnt) & "")  ' RA_RST3
            pAry(LneCnt, 33) = IIf(arrRecords2(45, PrtCnt) = 0, "-", arrRecords2(45, PrtCnt) & "")  ' RA_RST_AVE
    '        pAry(LneCnt, 33) = IIf(IsNull(arrRecords1(25, 0)) = True, "-", "OK" & "") '  UT
            pAry(LneCnt, 34) = IIf(arrRecords1(25, 0) = " ", "-", "OK" & "") '  UT
            
    '        �����ʾ׷�ӵ�����Ҫ��  ��ѧ�� 20120221
            
                
            If sADD_MATR = "A" Then
                        pAry(LneCnt, 35) = IIf(arrRecords2(48, PrtCnt) = 0, "-", arrRecords2(48, PrtCnt) & "")    '�����ʱ���׷�Ӵ�ӡӲ��
    '        ElseIf sADD_MATR = "B" Then
    '                    pAry(LneCnt, 35) = IIf(arrRecords2(49, PrtCnt) = 0, "-", arrRecords2(49, PrtCnt) & "")    '  �����ӡ
            Else
                        pAry(LneCnt, 35) = "-"                          ' �հ�-
            End If
        
    '        pAry(LneCnt, 35) = IIf(arrRecords1(28, 0) = "A", arrRecords2(48, PrtCnt) & "", "-")     ' HARD_RST
        End If
        
        
        
'20100611 sun bin  end
        
        LneCnt_D = LneCnt_D + 1
        pAry1(LneCnt_D, 1) = arrRecords2(1, PrtCnt) & ""                 ' PROD_NAME
        pAry1(LneCnt_D, 2) = arrRecords2(29, PrtCnt) & ""                ' NOTE
        

        lSumQNTY = lSumQNTY + arrRecords2(3, PrtCnt)                    'PAGE SUM QUANTITY
        dSumWGT = dSumWGT + arrRecords2(4, PrtCnt)                      'PAGE SUM WEIGHT
        dSumWGT = IIf(dSumWGT < 1, "0" & dSumWGT, dSumWGT)              '��ֹ����С��1t ����.XXX ���Ұ�ģ���Ϊ����

        lSumQNTY_T = lSumQNTY_T + arrRecords2(3, PrtCnt)                'TOTAL SUM QUANTITY
        dSumWGT_T = dSumWGT_T + arrRecords2(4, PrtCnt)                  'TOTAL SUM WEIGHT
        dSumWGT_T = IIf(dSumWGT_T < 1, "0" & dSumWGT_T, dSumWGT_T)      '��ֹ����С��1t ����.XXX ���Ұ�ģ���Ϊ����
        
        If LneCnt = 5 Then
            On Error GoTo ErrProc
            
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
            
            '����˵��������AQD053Cģ��
            If Report_KND = "T" Then
                xlApp.Workbooks.Open (App.Path & "\AQD053C.xls")
            Else
                If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
                 If iSave_State = 3 Then
                  xlApp.Workbooks.Open (App.Path & "\AQD051A.xls")
                 Else
                    xlApp.Workbooks.Open (App.Path & "\AQD051C.xls")
                 End If
                ElseIf UCase(Mid(arrRecords1(8, 0), 1, 3)) = "NV" Then
                    xlApp.Workbooks.Open (App.Path & "\AQD052C.xls")
                Else
                    xlApp.Workbooks.Open (App.Path & "\AQD050C.xls")
                End If
            End If
            
            Set xlSheet = xlApp.Worksheets("Sheet1")
              
            Call MillSheetPrint_S_Head(arrRecords1)
                        
            xlSheet.Range("D23").Value = lSumQNTY         'PAGE_CNT
            xlSheet.Range("F23").Value = dSumWGT          'PAGE_WGT
                                   

            If PrtCnt = RowCNT Then
                xlSheet.Range("B42").Value = "�ϼ�(Total)"
                xlSheet.Range("B43").Value = lSumQNTY_T & " Piece"        'SUM_CNT
                xlSheet.Range("C43").Value = IIf(dSumWGT_T < 1, "0" & dSumWGT_T, dSumWGT_T) & " ton"        'SUM_WGT
            End If
            
            xlSheet.Range("P44").Value = Round((RowCNT + 2) / 3, 0) & " - " & Round((PrtCnt + 2) / 3, 0)
            Page_no = Round((PrtCnt + 2) / 3, 0)
'��һ��������
            If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
               xlSheet.Range("B17:AL22").Value = pAry
            Else
               xlSheet.Range("B17:AJ22").Value = pAry
            End If

            xlSheet.Range("B25").Value = pAry1(1, 1)
            xlSheet.Range("E24").Value = Left(pAry1(1, 2), 150)
            
            If Len(pAry1(1, 2)) > 150 Then
                xlSheet.Range("C25").Value = Mid(pAry1(1, 2), 151, 165)
            End If
            
            If Len(pAry1(1, 2)) > 315 Then
                xlSheet.Range("C26").Value = Mid(pAry1(1, 2), 316, 165)
            End If
            If Len(pAry1(1, 2)) > 480 Then
                xlSheet.Range("C27").Value = Mid(pAry1(1, 2), 481, Len(pAry1(1, 2)))
'                xlSheet.Range("C24").Value = Mid(pAry1(1, 2), 331, Len(pAry1(1, 2)) - 331)
            End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
'�ڶ���������
            If arrRecords1(8, 0) = "NV" And Trim(pAry1(2, 1)) = "" Then  'NV�����š���Ʒ��Ϊ�����ӡ��-�� 20110812 liuxiang
                xlSheet.Range("B29").Value = "-"
                xlSheet.Range("E28").Value = "-"
            Else
                xlSheet.Range("B29").Value = pAry1(2, 1)
                xlSheet.Range("E28").Value = Left(pAry1(2, 2), 150)
            End If
            
            If Len(pAry1(2, 2)) > 150 Then
                xlSheet.Range("C29").Value = Mid(pAry1(2, 2), 151, 165)
            End If
            
            If Len(pAry1(2, 2)) > 315 Then
                xlSheet.Range("C30").Value = Mid(pAry1(2, 2), 316, 165)
            End If
            
            If Len(pAry1(2, 2)) > 480 Then
                xlSheet.Range("C31").Value = Mid(pAry1(2, 2), 481, Len(pAry1(2, 2)))
'                xlSheet.Range("C28").Value = Mid(pAry1(2, 2), 331, Len(pAry1(2, 2)) - 331)
            End If
'������������
            If arrRecords1(8, 0) = "NV" And Trim(pAry1(3, 1)) = "" Then  'NV�����š���Ʒ��Ϊ�����ӡ��-�� 20110812 liuxiang
                xlSheet.Range("B33").Value = "-"
                xlSheet.Range("E32").Value = "-"
            Else
                xlSheet.Range("B33").Value = pAry1(3, 1)
                xlSheet.Range("E32").Value = Left(pAry1(3, 2), 150)
            End If
            
             If Len(pAry1(3, 2)) > 150 Then
                xlSheet.Range("C33").Value = Mid(pAry1(3, 2), 151, 165)
            End If
            
            If Len(pAry1(3, 2)) > 315 Then
                xlSheet.Range("C34").Value = Mid(pAry1(3, 2), 316, 165)
            End If
            If Len(pAry1(3, 2)) > 480 Then
                xlSheet.Range("C35").Value = Mid(pAry1(3, 2), 481, Len(pAry1(3, 2)))
'                xlSheet.Range("C32").Value = Mid(pAry1(3, 2), 331, Len(pAry1(3, 2)) - 331)
            End If
            
            xlSheet.Range("H42").Value = sUserName
'           Worksheets("Sheet1").Columns("C").Hidden = True
            
'            xlSheet.Columns("S").Hidden = True
            
            Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, Save_State, Save_Path)
            If Save_State = 0 Or Save_State = 1 Or Save_State = 3 Then
'           xlApp.Application.Visible = True
            
                If Report_KND = "T" Then        ' T��˵����
                    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
                Else
                    If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "CCS" Or UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Or Save_State = 3 Then
                        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
                    Else
                        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
                    End If
                End If

            End If

            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit
            
            LneCnt = -1
            lSumQNTY = 0
            dSumWGT = 0
            LneCnt_D = 0
            ReDim pAry(1 To 6, 1 To 38)
            ReDim pAry1(1 To 6, 1 To 2)
            
        End If

    Loop Until PrtCnt = RowCNT
    
    If LneCnt <> -1 Then
    
        On Error GoTo ErrProc
        
        Set xlApp = GetObject("", "Excel.Application")
        If Err.Number = 429 Then
            Set xlApp = CreateObject("", "Excel.Application")
        End If

         '����˵��������AQD053Cģ��
        If Report_KND = "T" Then
                xlApp.Workbooks.Open (App.Path & "\AQD053C.xls")
        Else
            If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
                If iSave_State = 3 Then
                  xlApp.Workbooks.Open (App.Path & "\AQD051A.xls")
                 Else
                    xlApp.Workbooks.Open (App.Path & "\AQD051C.xls")
                 End If
            ElseIf UCase(Mid(arrRecords1(8, 0), 1, 3)) = "NV" Then
                xlApp.Workbooks.Open (App.Path & "\AQD052C.xls")
            Else
                xlApp.Workbooks.Open (App.Path & "\AQD050C.xls")
            End If
        End If
        Set xlSheet = xlApp.Worksheets("Sheet1")
 'SUN BIN 2010.01.30 END
        
        Call MillSheetPrint_S_Head(arrRecords1)
        
        xlSheet.Range("D23").Value = lSumQNTY                     'SUM_CNT
        xlSheet.Range("F23").Value = dSumWGT                      'SUM_WGT
        
        xlSheet.Range("B42").Value = "�ϼ�(Total)"
        xlSheet.Range("B43").Value = lSumQNTY_T & " Piece"                   'SUM_CNT
        xlSheet.Range("C43").Value = IIf(dSumWGT_T < 1, "0" & dSumWGT_T, dSumWGT_T) & " ton"                   'SUM_WGT
        xlSheet.Range("P44").Value = Round((RowCNT + 2) / 3, 0) & " - " & Round((PrtCnt + 2) / 3, 0)
        Page_no = Round((PrtCnt + 2) / 3, 0)

'��һ��������

        If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
          xlSheet.Range("B17:AL22").Value = pAry
        Else
          xlSheet.Range("B17:AJ22").Value = pAry
        End If
        'xlSheet.Range("B17:AJ22").Value = pAry
        xlSheet.Range("B25").Value = pAry1(1, 1)
        xlSheet.Range("E24").Value = Left(pAry1(1, 2), 150)
            
            If Len(pAry1(1, 2)) > 150 Then
                xlSheet.Range("C25").Value = Mid(pAry1(1, 2), 151, 165)
            End If
            
            If Len(pAry1(1, 2)) > 315 Then
                xlSheet.Range("C26").Value = Mid(pAry1(1, 2), 316, 165)
            End If
            If Len(pAry1(1, 2)) > 480 Then
                xlSheet.Range("C27").Value = Mid(pAry1(1, 2), 481, Len(pAry1(1, 2)))
'                xlSheet.Range("C24").Value = Mid(pAry1(1, 2), 331, Len(pAry1(1, 2)) - 331)
            End If
        

'�ڶ���������
    
        If arrRecords1(8, 0) = "NV" And Trim(pAry1(2, 1)) = "" Then    'NV�����š���Ʒ��Ϊ�����ӡ��-�� 20110812 liuxiang
            xlSheet.Range("B29").Value = "-"
            xlSheet.Range("E28").Value = "-"
        Else
            xlSheet.Range("B29").Value = pAry1(2, 1)
            xlSheet.Range("E28").Value = Left(pAry1(2, 2), 150)
        End If
            
            If Len(pAry1(2, 2)) > 150 Then
                xlSheet.Range("C29").Value = Mid(pAry1(2, 2), 151, 165)
            End If
            
            If Len(pAry1(2, 2)) > 315 Then
                xlSheet.Range("C30").Value = Mid(pAry1(2, 2), 316, 165)
            End If
            
            If Len(pAry1(2, 2)) > 480 Then
                xlSheet.Range("C31").Value = Mid(pAry1(2, 2), 481, Len(pAry1(2, 2)))
'                xlSheet.Range("C28").Value = Mid(pAry1(2, 2), 331, Len(pAry1(2, 2)) - 331)
            End If
'������������

        If arrRecords1(8, 0) = "NV" And Trim(pAry1(3, 1)) = "" Then    'NV�����š���Ʒ��Ϊ�����ӡ��-�� 20110812 liuxiang
            xlSheet.Range("B33").Value = "-"
            xlSheet.Range("E32").Value = "-"
        Else
            xlSheet.Range("B33").Value = pAry1(3, 1)
            xlSheet.Range("E32").Value = Left(pAry1(3, 2), 150)
        End If
            
             If Len(pAry1(3, 2)) > 150 Then
                xlSheet.Range("C33").Value = Mid(pAry1(3, 2), 151, 165)
            End If
            
            If Len(pAry1(3, 2)) > 315 Then
                xlSheet.Range("C34").Value = Mid(pAry1(3, 2), 316, 165)
            End If
            If Len(pAry1(3, 2)) > 480 Then
                xlSheet.Range("C35").Value = Mid(pAry1(3, 2), 481, Len(pAry1(3, 2)))
'                xlSheet.Range("C32").Value = Mid(pAry1(3, 2), 331, Len(pAry1(3, 2)) - 331)
            End If

        xlSheet.Range("H42").Value = sUserName
'
'        xlSheet.Columns("S").Hidden = True
               
        Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, Save_State, Save_Path)
        If Save_State = 0 Or Save_State = 1 Or Save_State = 3 Then
'           xlApp.Application.Visible = True
            If Report_KND = "T" Then
                xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Else
                If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "CCS" Or UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Or Save_State = 3 Then
                    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
                Else
                    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
                End If
            End If
        End If

        Set xlSheet = Nothing
        xlApp.ActiveWorkbook.Close False
        xlApp.Quit
        
    End If
            
    Set xlApp = Nothing
    
    Exit Function
    
ErrProc:
    If Err.Number = 429 Then
        MsgBox "Microsoft Excel Program Not Installed"
    Else
        MsgBox Err.Number & Err.Description
    End If
    MillSheetPrint_S = "ERROR"
    
    Set xlSheet = Nothing
    xlApp.ActiveWorkbook.Close False
    xlApp.Quit
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
    
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_S_Head
'   2.Name         : Ship certificate print(Head table)
'   3.Input  Value : arrRecords1
'   4.Return Value :
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Sub MillSheetPrint_S_Head(ByVal arrRecords1 As Variant)
    Dim sDate As String
    Dim sCert_No As String
    Dim sPROD_CD As String
    Dim sSTDSPEC As String
    Dim sNEWSPEC As String  '���ڸ��ݶ�����ע�ж��Ƿ�ִ���´���   ���� 2011.10.9
    Dim sADD_MATR As String  '׷�������ʱ���ӡ Ӳ�ȡ����� Ӳ�����ȣ���������
    
    '�����˵���飬һ�����ݲ���ӡ
    If Report_KND = "T" Then
        arrRecords1(0, 0) = ""      '�ʱ����
        arrRecords1(1, 0) = ""      '�������ɺ�
        arrRecords1(12, 0) = ""     '���ƺ�
    End If
    
    sDate = GetPrintDate()
    sCert_No = arrRecords1(0, 0)
    sPROD_CD = Mid(sCert_No, 1, 2)
    sNEWSPEC = IIf(InStr(arrRecords1(27, 0), "ִ���¹淶") = 0, "0", "1") '"0"���ɴ���  "1"���´���
    
    'xlSheet.Range("W2").Value = arrRecords1(23, 0) & ""       'QLTY_REC_NO
    If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
      xlSheet.Range("AG4").Value = arrRecords1(0, 0) & ""        'CERT_NO
    Else
      xlSheet.Range("AE4").Value = arrRecords1(0, 0) & ""        'CERT_NO
    End If
    
    xlSheet.Range("C2").Value = arrRecords1(1, 0) & ""        'PROD_SPEC_NO
    
    If arrRecords1(8, 0) = "NV" Then
'       sSTDSPEC = "DNV Rules Pt2. Ch2. 2011"                     '2011.8.12  liuxiang
       sSTDSPEC = "DNV Rules Pt2."                     '2011.8.12  liuxiang
       xlSheet.Range("C3").Value = arrRecords1(1, 0) & ""        'PROD_SPEC_NO
    ElseIf arrRecords1(8, 0) = "GB" Then
       sSTDSPEC = "GB712" & " �淶"
'    wch 20130216
    Else
       sSTDSPEC = arrRecords1(8, 0) & " �淶"
    End If
    
    If InStr(arrRecords1(2, 0), "DNV-OS-B101") <> 0 Then
       sSTDSPEC = "DNV-OS-B101"
    End If
    
'    xlSheet.Range("C5").Value = sSTDSPEC 'STDSPEC_NAME

        
'    xlSheet.Range("C9").Value = arrRecords1(3, 0) & ""        'COND_SUPPLY
    
    '�����NV���ƺ�ǰ�ӡ�NV �� 2012.5.23 ����
    If arrRecords1(8, 0) = "NV" Then
        xlSheet.Range("E11").Value = "NV " + arrRecords1(4, 0) + ")" + " ��������ߴ磨mm):" + arrRecords1(5, 0)          'PROD_NAME1
        xlSheet.Range("G12").Value = "NV " + arrRecords1(4, 0) + " Dimensions of test specimens:" + arrRecords1(5, 0) + ")" 'PROD_NAME2
        xlSheet.Range("C5").Value = sSTDSPEC 'STDSPEC_NAME
        xlSheet.Range("C9").Value = arrRecords1(3, 0) & ""        'COND_SUPPLY
    ElseIf arrRecords1(8, 0) = "LR" Then
        xlSheet.Range("E11").Value = arrRecords1(4, 0) + ")" + " ��������ߴ磨mm):" + arrRecords1(5, 0)            'PROD_NAME1
        xlSheet.Range("G12").Value = arrRecords1(4, 0) + " Dimensions of test specimens:" + arrRecords1(5, 0) + ")"  'PROD_NAME2
        xlSheet.Range("C5").Value = sSTDSPEC 'STDSPEC_NAME
        xlSheet.Range("C8").Value = arrRecords1(3, 0) & ""        'COND_SUPPLY
'    ElseIf arrRecords1(8, 0) = "RS" Then ' RS �� PC Guhf
'        xlSheet.Range("E11").Value = "PC" + arrRecords1(4, 0) + ")" + " ��������ߴ磨mm):" + arrRecords1(5, 0)          'PROD_NAME1
'        xlSheet.Range("G12").Value = arrRecords1(4, 0) + " Dimensions of test specimens:" + arrRecords1(5, 0) + ")"  'PROD_NAME2
'        xlSheet.Range("C6").Value = sSTDSPEC 'STDSPEC_NAME
'        xlSheet.Range("C9").Value = arrRecords1(3, 0) & ""        'COND_SUPPLY
    Else
        If arrRecords1(8, 0) = "ABS" And Left(arrRecords1(4, 0), 1) = "D" And (arrRecords1(3, 0) = "CR" Or arrRecords1(3, 0) = "TMCP") Then
        
        xlSheet.Range("E11").Value = Replace(arrRecords1(4, 0), "N", "") + ")" + " ��������ߴ磨mm):" + arrRecords1(5, 0)          'PROD_NAME1
        xlSheet.Range("G12").Value = Replace(arrRecords1(4, 0), "N", "") + " Dimensions of test specimens:" + arrRecords1(5, 0) + ")"  'PROD_NAME2
        Else
        xlSheet.Range("E11").Value = arrRecords1(4, 0) + ")" + " ��������ߴ磨mm):" + arrRecords1(5, 0)            'PROD_NAME1
        xlSheet.Range("G12").Value = arrRecords1(4, 0) + " Dimensions of test specimens:" + arrRecords1(5, 0) + ")"  'PROD_NAME2
        End If
        
        xlSheet.Range("C6").Value = sSTDSPEC 'STDSPEC_NAME
        xlSheet.Range("C9").Value = arrRecords1(3, 0) & ""        'COND_SUPPLY
    End If
    
'�ŵ��ʱ��������Ϣ�� ��ѧ��  2011 06 23
'    xlSheet.Range("H13").Value = arrRecords1(24, 0) & ""        'PROD_SIZE
'  --WCH 2013-01-11 CCS�ʱ��鲻ͬ��ʽ
   If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
     xlSheet.Range("AG6").Value = Left(sDate, 4) + "��" + Mid(sDate, 6, 2) + "��" + Mid(sDate, 9, 2) + "��"          'CERT_RPT_DATE1
     xlSheet.Range("AG7").Value = Mid(sDate, 6, 2) + "/" + Mid(sDate, 9, 2) + "/" + Left(sDate, 4)          'CERT_RPT_DATE2
   Else
     xlSheet.Range("AE6").Value = Left(sDate, 4) + "��" + Mid(sDate, 6, 2) + "��" + Mid(sDate, 9, 2) + "��"          'CERT_RPT_DATE1
     If arrRecords1(8, 0) = "NV" Then
       xlSheet.Range("AE7").Value = Left(sDate, 4) + "-" + Mid(sDate, 6, 2) + "-" + Mid(sDate, 9, 2)
     Else
       xlSheet.Range("AE7").Value = Mid(sDate, 6, 2) + "/" + Mid(sDate, 9, 2) + "/" + Left(sDate, 4)          'CERT_RPT_DATE2
    End If
   End If
'    xlSheet.Range("W37").Value = IIf(arrRecords1(8, 0) = "NV", "DNV", arrRecords1(8, 0)) & ""       'STD_ORGAN
    xlSheet.Range("H42").Value = arrRecords1(9, 0) & ""       'TEST_EMP
    xlSheet.Range("C13").Value = arrRecords1(12, 0) & ""      'CONTROL
    xlSheet.Range("B37").Value = arrRecords1(13, 0) & ""      'TEXT1
'20100611 sun bin start
'If arrRecords1(8, 0) = "NV" And Trim(arrRecords1(25, 0)) = "" Then
'    xlSheet.Range("I13").Value = "-" & ""                           'UST_STD ���NV��̽�˴�ӡ��-�� 20110812 liuxiang
'Else
'    xlSheet.Range("I13").Value = arrRecords1(25, 0) & ""            'UST_STD
'End If
If Trim(arrRecords1(25, 0)) = "" Then
    xlSheet.Range("I13").Value = "-" & ""                           'UST_STD ���NV��̽�˴�ӡ��-�� 20110812 liuxiang
Else
    xlSheet.Range("I13").Value = arrRecords1(25, 0) & ""            'UST_STD
End If

    'xlSheet.Range("V13").Value = arrRecords1(26, 0) & ""            '�ͻ�����
    xlSheet.Range("V13").Value = arrRecords1(26, 0) + Chr(10) + "�����ţ�" + arrRecords1(33, 0) + "-" + arrRecords1(34, 0) & ""      '�ͻ�����
'20100611 sun bin end
    If arrRecords1(8, 0) = "ABS" Then
        xlSheet.Range("B38").Value = "2.The mark of AB/" + arrRecords1(4, 0) + " is stamped on the end of each plate."    'TEXT2
    Else
        xlSheet.Range("B38").Value = arrRecords1(14, 0) & ""      'TEXT2
    End If
    If arrRecords1(8, 0) = "CCS" Or arrRecords1(8, 0) = "RINA" Then
        If arrRecords1(8, 0) = "CCS" Then
        xlSheet.Range("B39").Value = arrRecords1(15, 0) & ""                                'TEXT3 CCS ����CEQ��ʽ
'        xlSheet.Range("B40").Value = arrRecords1(16, 0) & arrRecords1(4, 0) & " steel."     'TEXT4
        xlSheet.Range("B40").Value = arrRecords1(16, 0) & ""                                 'TEXT4
        Else
        xlSheet.Range("B39").Value = arrRecords1(30, 0) & ""                                'TEXT3 �ĳ� ���ݶ����ɷݱ�׼�Ĺ�ʽ ���ץȡ��ʽ ��ѧ�� 20120224
        xlSheet.Range("B40").Value = arrRecords1(16, 0) & arrRecords1(4, 0) & " steel."     'TEXT4
        End If
    Else
        
        xlSheet.Range("B39").Value = arrRecords1(30, 0) & ""                                'TEXT3 �ĳ� ���ݶ����ɷݱ�׼�Ĺ�ʽ ���ץȡ��ʽ ��ѧ�� 20120224
        xlSheet.Range("B40").Value = arrRecords1(16, 0) & ""                                'TEXT4
    End If
'    xlSheet.Range("B37").Value = arrRecords1(16, 0) & ""      'TEXT4
    If arrRecords1(8, 0) = "CCS" Then
        xlSheet.Range("B41").Value = arrRecords1(17, 0) & ""
        xlSheet.Range("G37").Value = arrRecords1(18, 0) & ""     'TEXT6
        xlSheet.Range("G38").Value = arrRecords1(19, 0) & ""      'TEXT7
        xlSheet.Range("G39").Value = arrRecords1(20, 0) & ""     'TEXT8
        xlSheet.Range("G40").Value = arrRecords1(21, 0) & ""     'TEXT9
        xlSheet.Range("G41").Value = arrRecords1(22, 0) & ""     'TEXT9
        xlSheet.Range("Y37").Value = arrRecords1(8, 0) & ""       'STD_ORGAN    --WCH 2013-01-11
    Else
'        xlSheet.Range("B41").Value = IIf(sNEWSPEC = "1", "", arrRecords1(17, 0) & "")  'TEXT5  ����sNEWSPEC�ж��Ƿ��ӡ
        xlSheet.Range("G37").Value = "                                       " + arrRecords1(18, 0) & ""    'TEXT6
        
        If InStr(arrRecords1(2, 0), "DNV-OS-B101") = 0 Then
        xlSheet.Range("G38").Value = "                                       " + arrRecords1(19, 0) & ""    'TEXT7
        xlSheet.Range("G39").Value = "                                       " + arrRecords1(20, 0) & ""    'TEXT8
        xlSheet.Range("G40").Value = "                                       " + arrRecords1(21, 0) + " " + IIf(arrRecords1(8, 0) = "NV", "DNV(R-1693).", arrRecords1(8, 0))     'TEXT9
        Else
        xlSheet.Range("G38").Value = "                                       " + "made by an approved process and has been"
        xlSheet.Range("G39").Value = "                                       " + "satisfactorily tested in accordance with DNV"
        xlSheet.Range("G40").Value = "                                       " + "Offshore Standards (R-1693)."
        End If
'        xlSheet.Range("G41").Value = IIf(sNEWSPEC = "1", "", arrRecords1(22, 0) & "") 'TEXT10 ����sNEWSPEC�ж��Ƿ��ӡ
'        xlSheet.Range("W37").Value = IIf(arrRecords1(8, 0) = "NV", "DNV", arrRecords1(8, 0)) & ""       'STD_ORGAN
        If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
         xlSheet.Range("AA37").Value = IIf(arrRecords1(8, 0) = "NV", "DNV", arrRecords1(8, 0)) & ""       'STD_ORGAN
        Else
         xlSheet.Range("Y37").Value = IIf(arrRecords1(8, 0) = "NV", "DNV", arrRecords1(8, 0)) & ""       'STD_ORGAN
        End If
    End If
    
'    xlSheet.Range("K37").Value = arrRecords1(18, 0) & ""      'TEXT6
'    xlSheet.Range("K38").Value = arrRecords1(19, 0) & ""      'TEXT7
'    xlSheet.Range("K39").Value = arrRecords1(20, 0) & ""      'TEXT8
'    xlSheet.Range("K40").Value = arrRecords1(21, 0) + " " + IIf(arrRecords1(8, 0) = "NV", "DNV", arrRecords1(8, 0)) + "."      'TEXT9
'    xlSheet.Range("K41").Value = IIf(sNEWSPEC = "1", "", arrRecords1(22, 0) & "") 'TEXT10 ����sNEWSPEC�ж��Ƿ��ӡ
    
'    If arrRecords1(29, 0) = "A" Then
'        sADD_MATR = "B"                        '  �����ӡ
'    End If
    If arrRecords1(28, 0) = "A" Then
        sADD_MATR = "A"                       '   Ӳ��
    End If
        
    If sADD_MATR = "A" Then
      If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
                xlSheet.Range("AJ14").Value = "Ӳ��" & Chr(10) & "HRC"   '�����ʱ���׷�Ӵ�ӡӲ��
      Else
               xlSheet.Range("AL14").Value = "Ӳ��" & Chr(10) & "HRC"   '�����ʱ���׷�Ӵ�ӡӲ��
      End If
'    ElseIf sADD_MATR = "B" Then
'                xlSheet.Range("AJ14").Value = "����" & Chr(10) & "B.D."    '  �����ӡ
    Else
        If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
                xlSheet.Range("AJ14").Value = "-"                          ' �հ�-
        Else
                xlSheet.Range("AL14").Value = "-"                          ' �հ�-
        End If
    End If
    
        
'    '������Ϣ 2014.7.24 ����
'    If arrRecords1(32, 0) = "Y" Then
'       xlSheet.Range("T12").Value = "������Ϣ"
'       If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
'          xlSheet.Range("Y12").Value = arrRecords1(31, 0) & ""
'       Else
'          xlSheet.Range("W12").Value = arrRecords1(31, 0) & ""
'       End If
'    End If
    
    '�в���������ʾ������Ϣ 2015.1.28 ����
    '��Ϊȡ������Ϣ������� 2015.8.31 ����
    If Trim(arrRecords1(31, 0)) <> "" Then
       xlSheet.Range("T12").Value = "������Ϣ:" + arrRecords1(31, 0) & ""
'       If UCase(Mid(arrRecords1(8, 0), 1, 3)) = "ABS" Then
'          xlSheet.Range("Y12").Value = arrRecords1(31, 0) & ""
'       Else
'          xlSheet.Range("W12").Value = arrRecords1(31, 0) & ""
'       End If
    End If
    
    
   xlSheet.Range("B41").Value = "ACCORDING TO EN10204:2004 3.2."
    
'    If UCase(sPROD_CD) = "HC" Then
'        xlSheet.Range("B14").Value = "��Ʒ��" + Chr$(10) + "Coil ��"
'        xlSheet.Range("B22").Value = "��Ʒ��" + Chr$(10) + "Coil ��"
'    Else
'        xlSheet.Range("B14").Value = "��Ʒ��" + Chr$(10) + "Plate ��"
'        xlSheet.Range("B22").Value = "��Ʒ��" + Chr$(10) + "Plate ��"
'    End If

End Sub

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - funGetQuery_P
'   2.Name         : Pipe certificate
'   3.Input  Value : sCertNo , iSave_State ,sSave_Path
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function funGetQuery_P(sCertNo As String, iSave_State As Integer, sSave_Path As String) As String
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
    Dim arrRecords3 As Variant
    Dim AdoRs As adodb.Recordset
    Dim sPROD_CD As String
    Dim sTable_PROD As String
    Dim sFieldName_NO As String
    
    sPROD_CD = Mid(sCertNo, 1, 2)
    If UCase(sPROD_CD) = "HC" Then
        sTable_PROD = "GP_COIL"
        sFieldName_NO = "COIL_NO"
    Else
        sTable_PROD = "GP_PLATE"
        sFieldName_NO = "PLATE_NO"
    End If
       
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT CERT_NO , STDSPEC_NAME , DECODE(TRIM(GF_ENDCUSTER_FIND(SHIP_ISP_NO)),'',GF_CUST_NAME(CUST_CD,''),GF_ENDCUSTER_FIND(SHIP_ISP_NO)) "
    sQuery = sQuery + ",QLTY_REC_NO, GF_PONO_FIND(ORD_NO),SHIP_ISP_NO,GF_EMPNAMEFIND(TEST_EMP) AS TEST_EMP,AQD0060C.F_SUM_CNT(CERT_NO) AS SUM_CNT, AQD0060C.F_SUM_WGT(CERT_NO) AS SUM_WGT"
    sQuery = sQuery + ",PROD_NAME, COND_SUPPLY, TRAIN_LINE_NAME,GF_STDSPEC_NAME_ENG(STDSPEC_STLGRD),IMPACT_SMP_SIZE"
    sQuery = sQuery + ",DECODE(Gf_ComnNameFind('Q0046',UST_FL) ,'ASTM A 435 / ASME SA-435','ASTM A435 / A435M-90','JB4730 J11'"
    sQuery = sQuery + ",'JB4730-94 ��','JB4730 J21','JB4730-94 ��','GB/T 2970 K11','GB/T 2970 ��','GB/T 2970 K21','GB/T 2970 ��','NO UST',' '"
    sQuery = sQuery + ",Gf_ComnNameFind('Q0046',UST_FL))"
    sQuery = sQuery + " FROM QP_CERT_HEAD WHERE CERT_NO  = '" & sCertNo & "'"
            
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_P = "Err DataBase"
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = "SELECT CERT_NO ,GF_MARKING_NO(PROD_NO) , GF_STLGRD_DETAIL(STLGRD), PROD_SIZE , PRDT_QNTY , PRDT_WGT"
    sQuery = sQuery + ", DECODE(C_RST,NULL,0,C_RST) ,  DECODE(MN_RST,NULL,0,MN_RST) , DECODE(P_RST,NULL,0,P_RST),DECODE(S_RST,NULL,0,S_RST) "
    sQuery = sQuery + ", DECODE(SI_RST,NULL,0,SI_RST) ,DECODE(CU_RST,NULL,0,CU_RST) , DECODE(NI_RST,NULL,0,NI_RST), DECODE(CR_RST,NULL,0,CR_RST)"
    sQuery = sQuery + ", DECODE(MO_RST,NULL,0,MO_RST) ,  DECODE(V_RST,NULL,0,V_RST),DECODE(TI_RST,NULL,0,TI_RST), DECODE(NB_RST,NULL,0,NB_RST) "
    'sQuery = sQuery + ", DECODE(GF_AQD0060C_FIND('Alt',PROD_NO,CERT_NO),0,GF_AQD0060C_FIND('Als',PROD_NO,CERT_NO),GF_AQD0060C_FIND('Alt',PROD_NO,CERT_NO)) "
    sQuery = sQuery + ", GF_AQD0060C_FIND('N',PROD_NO,CERT_NO),DECODE(CEQ_RST,NULL,0,CEQ_RST),GF_AQD0060C_FIND('Pcm',PROD_NO,CERT_NO)"
    sQuery = sQuery + ", DECODE(GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Alt'),NULL,0,GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Alt')),"
    sQuery = sQuery + " DECODE(GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Ca'),NULL,0,GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Ca')), "
    sQuery = sQuery + " DECODE(GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Als'),NULL,0,GF_CHEM_RSLT(SUBSTR(PROD_NO,1,8),'Als')) "
    sQuery = sQuery + " FROM QP_CERT_DETAIL WHERE CERT_NO  = '" & sCertNo & "'"
            
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_P = "Err DataBase"
        Exit Function
    End If
    arrRecords2 = AdoRs.GetRows
    AdoRs.Close

    sQuery = "SELECT CERT_NO ,GF_MARKING_NO(PROD_NO) "
    sQuery = sQuery + ", DECODE(YP_RST,NULL,0,YP_RST), DECODE(TS_RST,NULL,0,TS_RST),DECODE(EL_RST,NULL,0,EL_RST),DECODE(YR_RST,NULL,0,YR_RST)"
    sQuery = sQuery + ", DECODE(BEND_RST,'Y','OK','-') , DECODE(UST_GRD_RST,NULL,'','OK'),DECODE(IMPACT_TMP,NULL,0,IMPACT_TMP), DECODE(IMPACT_RST1,NULL,0,IMPACT_RST1)"
    sQuery = sQuery + ", DECODE(IMPACT_RST2,NULL,0,IMPACT_RST2),DECODE(IMPACT_RST3,NULL,0,IMPACT_RST3),DECODE(IMPACT_RST4,NULL,0,IMPACT_RST4),DECODE(IMPACT_RST5,NULL,0,IMPACT_RST5)"
    sQuery = sQuery + ", DECODE(IMPACT_RST6,NULL,0,IMPACT_RST6),DECODE(IMPACT_RST_AVE,NULL,0,IMPACT_RST_AVE),DECODE(IMPACT_RATE_RST1,NULL,0,IMPACT_RATE_RST1)"
    sQuery = sQuery + ", DECODE(IMPACT_RATE_RST2,NULL,0,IMPACT_RATE_RST2),DECODE(IMPACT_RATE_RST3,NULL,0,IMPACT_RATE_RST3),DECODE(DWTT_TMP,NULL,0,DWTT_TMP)"
    sQuery = sQuery + ", DECODE(DWTT_YP_RST1,NULL,0,DWTT_YP_RST1),DECODE(DWTT_YP_RST2,NULL,0,DWTT_YP_RST2),DECODE(DWTT_YP_RST_AVE,NULL,0,DWTT_YP_RST_AVE)"
    sQuery = sQuery + ",GF_GET_HARDRSLT(PROD_NO,'PP'),0,0"
    sQuery = sQuery + ", DECODE(GRAIN_SIZE_RST,NULL,0,GRAIN_SIZE_RST),DECODE(NON_METAL_DSC_RST,'Y','OK','N','NO'),BELT_STR_DSC_RST" ','Y','�ϸ�','N','���ϸ�')"
    sQuery = sQuery + " FROM QP_CERT_DETAIL WHERE CERT_NO  = '" & sCertNo & "'"
              
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        funGetQuery_P = "Err DataBase"
        Exit Function
    End If
    arrRecords3 = AdoRs.GetRows
    AdoRs.Close
    
    Set AdoRs = Nothing
    
    funGetQuery_P = MillSheetPrint_P(iSave_State, sSave_Path, arrRecords1, arrRecords2, arrRecords3)
    
End Function


'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_P
'   2.Name         : Pipe certificate print(Detail table)
'   3.Input  Value : iSave_State ,sSave_Path ,arrRecords1 ,arrRecords2
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 06 .27
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.2
'--------------------------------------------------------------------------------------------------------
Private Function MillSheetPrint_P(ByVal iSave_State As Integer, sSave_Path As String, ByVal arrRecords1 As Variant, ByVal arrRecords2 As Variant, ByVal arrRecords3 As Variant) As String
    Dim RowCNT      As Long
    Dim PrtCnt      As Long
    Dim LneCnt      As Long
    Dim pAry11()    As String
    Dim pAry12()    As String
    Dim pAry13()    As String
    Dim pAry14()    As String
    Dim pAry15()    As String
    Dim pAry2()     As String
    Dim iDx         As Integer
    Dim sRow        As String
    Dim lSumQNTY    As Long
    Dim dSumWGT     As Double
    Dim lSumQNTY_T    As Long
    Dim dSumWGT_T     As Double
    Dim Save_State  As Integer
    Dim Save_Path As String
    Dim Page_no As String
    Dim Cert_No         As String
    Dim NBVTI_SUM     As Double
    Dim CRCUNI_SUM     As Double
       
    Save_State = iSave_State
    Save_Path = sSave_Path
    Cert_No = arrRecords2(0, 0)
    
    If IsEmpty(arrRecords1) Or IsEmpty(arrRecords2) Or IsEmpty(arrRecords3) Then
        MillSheetPrint_P = "Err Data"
        Exit Function
    End If
    
    RowCNT = UBound(arrRecords2, 2)
    
    PrtCnt = -1
    LneCnt = 0
    lSumQNTY = 0
    dSumWGT = 0
    lSumQNTY_T = 0
    dSumWGT_T = 0
    
    ReDim pAry11(1 To 5, 1 To 2)
    ReDim pAry12(1 To 5, 1 To 1)
    ReDim pAry13(1 To 5, 1 To 1)
    ReDim pAry14(1 To 5, 1 To 1)
    ReDim pAry15(1 To 5, 1 To 16)
    ReDim pAry2(1 To 5, 1 To 24)

    Do
        
        LneCnt = LneCnt + 1
        PrtCnt = PrtCnt + 1

        pAry11(LneCnt, 1) = arrRecords2(1, PrtCnt) & ""                 ' PROD_NO
        '-----------------------------HJD-------------------------------------------------------
        If Trim(arrRecords2(2, PrtCnt) & "") = "X70HIC" Then
             pAry11(LneCnt, 2) = "X70"                 ' STLGRD
        
        ElseIf Trim(arrRecords2(2, PrtCnt) & "") = "X65-M" Then
             pAry11(LneCnt, 2) = "X65"                 ' STLGRD
        Else
            pAry11(LneCnt, 2) = arrRecords2(2, PrtCnt) & ""                 ' STLGRD
        
        End If

        'pAry11(LneCnt, 2) = arrRecords2(2, PrtCnt) & ""                 ' STLGRD
        pAry12(LneCnt, 1) = arrRecords2(3, PrtCnt) & ""                 ' PROD_SIZE
        pAry13(LneCnt, 1) = arrRecords2(4, PrtCnt) & ""                 ' QNTY
        pAry14(LneCnt, 1) = arrRecords2(5, PrtCnt) & ""                 ' WGT
        pAry15(LneCnt, 1) = IIf(Val(arrRecords2(6, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(6, PrtCnt) & "") * 100)                    ' C_RST
        
        pAry15(LneCnt, 2) = IIf(Val(arrRecords2(7, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(7, PrtCnt) & "") * 100)                    ' MN_RST
        
        pAry15(LneCnt, 3) = IIf(Val(arrRecords2(8, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(8, PrtCnt) & "") * 1000)                   ' P_RST
        
        pAry15(LneCnt, 4) = IIf(Val(arrRecords2(9, PrtCnt) & "") = 0, _
        "-", Val_zero(Val(arrRecords2(9, PrtCnt) & "") * 1000))                   ' S_RST
        
        pAry15(LneCnt, 5) = IIf(Val(arrRecords2(10, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(10, PrtCnt) & "") * 100)                   ' SI_RST
        
'        pAry15(LneCnt, 6) = IIf(Val(arrRecords2(11, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(11, PrtCnt) & "") * 100)                   ' CU_RST
'
'        pAry15(LneCnt, 7) = IIf(Val(arrRecords2(12, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(12, PrtCnt) & "") * 100)                   ' NI_RST
'
'        pAry15(LneCnt, 8) = IIf(Val(arrRecords2(13, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(13, PrtCnt) & "") * 100)                   ' CR_RST

'       pAry15(LneCnt, 6) = IIf(Val(arrRecords2(15, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(15, PrtCnt) & "") * 1000) + IIf(Val(arrRecords2(16, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(16, PrtCnt) & "") * 1000) + IIf(Val(arrRecords2(17, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(17, PrtCnt) & "") * 1000)                      ' V+Ti+Nb
        
        NBVTI_SUM = (Val(arrRecords2(15, PrtCnt) & "") * 1000) + (Val(arrRecords2(16, PrtCnt) & "") * 1000) + (Val(arrRecords2(17, PrtCnt) & "") * 1000)
                  
        pAry15(LneCnt, 6) = IIf(Val(NBVTI_SUM) = 0, "-", NBVTI_SUM)


        pAry15(LneCnt, 9) = IIf(Val(arrRecords2(14, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(14, PrtCnt) & "") * 100)                   ' MO_RST
        
'        If Mid(Trim(arrRecords1(4, 0)), 1, 10) = "06NGE/P084" Then
'            pAry15(LneCnt, 10) = IIf(Val(arrRecords2(22, PrtCnt) & "") = 0, _
'            "-", Val(arrRecords2(22, PrtCnt) & "") * 1000)                  ' CA_RST
'        Else
'            pAry15(LneCnt, 10) = IIf(Val(arrRecords2(15, PrtCnt) & "") = 0, _
'            "-", Val(arrRecords2(15, PrtCnt) & "") * 1000)                  ' V_RST
'        End If
'
'        pAry15(LneCnt, 11) = IIf(Val(arrRecords2(16, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(16, PrtCnt) & "") * 1000)                  ' TI_RST
'
'        pAry15(LneCnt, 12) = IIf(Val(arrRecords2(17, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(17, PrtCnt) & "") * 1000)                  ' NB_RST

'       pAry15(LneCnt, 10) = IIf(Val(arrRecords2(11, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(11, PrtCnt) & "") * 100) + IIf(Val(arrRecords2(12, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(12, PrtCnt) & "") * 100) + IIf(Val(arrRecords2(13, PrtCnt) & "") = 0, _
'        "-", Val(arrRecords2(13, PrtCnt) & "") * 100)                       'Cu+Ni+Cr
        
        CRCUNI_SUM = (Val(arrRecords2(11, PrtCnt) & "") * 100) + (Val(arrRecords2(12, PrtCnt) & "") * 100) + (Val(arrRecords2(13, PrtCnt) & "") * 100)
        pAry15(LneCnt, 10) = IIf(Val(CRCUNI_SUM) = 0, "-", CRCUNI_SUM)
        
        If Mid(Trim(arrRecords1(4, 0)), 1, 10) = "06NGE/P084" Then
            pAry15(LneCnt, 13) = IIf(Val(arrRecords2(23, PrtCnt) & "") = 0, _
            "-", Val(arrRecords2(23, PrtCnt) & "") * 1000)                  ' ALS_RST
        Else
            pAry15(LneCnt, 13) = IIf(Val(arrRecords2(21, PrtCnt) & "") = 0, _
            "-", Val(arrRecords2(21, PrtCnt) & "") * 1000)                  ' ALT_RST
        End If
        pAry15(LneCnt, 14) = IIf(Val(arrRecords2(18, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(18, PrtCnt) & "") & "")                    ' N_RST
        
        pAry15(LneCnt, 15) = IIf(Val(arrRecords2(19, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(19, PrtCnt) & "") * 100)                   ' CEQ_RST
        pAry15(LneCnt, 16) = IIf(Val(arrRecords2(20, PrtCnt) & "") = 0, _
        "-", Val(arrRecords2(20, PrtCnt) & "") * 100)                   ' PCM_RST
        
        pAry2(LneCnt, 1) = arrRecords3(1, PrtCnt) & ""                  ' PROD_NO
        pAry2(LneCnt, 2) = IIf(arrRecords3(2, PrtCnt) = 0, "-", arrRecords3(2, PrtCnt) & "") ' YP_RST
        pAry2(LneCnt, 3) = IIf(arrRecords3(3, PrtCnt) = 0, "-", arrRecords3(3, PrtCnt) & "") ' TS_RST
        pAry2(LneCnt, 4) = IIf(arrRecords3(4, PrtCnt) = 0, "-", arrRecords3(4, PrtCnt) & "") ' EL_RST
        pAry2(LneCnt, 5) = IIf(Val(arrRecords3(5, PrtCnt) & "") = 0, _
        "-", Val(arrRecords3(5, PrtCnt) & ""))                                                  ' YS_RST
        pAry2(LneCnt, 6) = IIf(arrRecords3(6, PrtCnt) = 0, "-", arrRecords3(6, PrtCnt) & "")    ' BEND_RST
        pAry2(LneCnt, 7) = IIf(arrRecords3(7, PrtCnt) = 0, "-", arrRecords3(7, PrtCnt) & "")    ' UST_GRD_RST
        pAry2(LneCnt, 8) = IIf(arrRecords3(9, PrtCnt) = 0, "-", arrRecords3(8, PrtCnt) & "")    ' IMPACT_TMP
        pAry2(LneCnt, 9) = IIf(arrRecords3(9, PrtCnt) = 0, "-", arrRecords3(9, PrtCnt) & "")    ' IMPACT_RST1
        pAry2(LneCnt, 10) = IIf(arrRecords3(10, PrtCnt) = 0, "-", arrRecords3(10, PrtCnt) & "") ' IMPACT_RST2
        pAry2(LneCnt, 11) = IIf(arrRecords3(11, PrtCnt) = 0, "-", arrRecords3(11, PrtCnt) & "") ' IMPACT_RST3
        
        If arrRecords3(12, PrtCnt) > 0 Or arrRecords3(13, PrtCnt) > 0 Or arrRecords3(14, PrtCnt) > 0 Then
            pAry2(LneCnt, 9) = IIf(arrRecords3(15, PrtCnt) = 0, "-", arrRecords3(15, PrtCnt) & "")    ' IMPACT_RST_AVE
            pAry2(LneCnt, 10) = IIf(arrRecords3(15, PrtCnt) = 0, "-", arrRecords3(15, PrtCnt) & "")     ' IMPACT_RST_AVE
            pAry2(LneCnt, 11) = IIf(arrRecords3(15, PrtCnt) = 0, "-", arrRecords3(15, PrtCnt) & "")     ' IMPACT_RST_AVE
        End If
        
        pAry2(LneCnt, 12) = IIf(arrRecords3(16, PrtCnt) = 0, "-", arrRecords3(16, PrtCnt) * 100 & "") ' IMPACT_RATE_RST1
        pAry2(LneCnt, 13) = IIf(arrRecords3(17, PrtCnt) = 0, "-", arrRecords3(17, PrtCnt) * 100 & "") ' IMPACT_RATE_RST2
        pAry2(LneCnt, 14) = IIf(arrRecords3(18, PrtCnt) = 0, "-", arrRecords3(18, PrtCnt) * 100 & "") ' IMPACT_RATE_RST3
        pAry2(LneCnt, 15) = IIf(arrRecords3(20, PrtCnt) = 0, "-", arrRecords3(19, PrtCnt) & "") ' DWTT_TMP
        pAry2(LneCnt, 16) = IIf(arrRecords3(20, PrtCnt) = 0, "-", arrRecords3(20, PrtCnt) & "")  ' DWTT_RST1
        pAry2(LneCnt, 17) = IIf(arrRecords3(21, PrtCnt) = 0, "-", arrRecords3(21, PrtCnt) & "")  ' DWTT_RST2
        pAry2(LneCnt, 18) = IIf(arrRecords3(22, PrtCnt) = 0, "-", arrRecords3(22, PrtCnt) & "") ' DWTT_AVE
        If arrRecords2(2, PrtCnt) = "5L B" Then                                                  ' HARD_RST JOMINY_RST1
            pAry2(LneCnt, 19) = "-"
        Else
            pAry2(LneCnt, 19) = IIf(arrRecords3(23, PrtCnt) = 0, "��230", arrRecords3(23, PrtCnt) & "")
        End If
'        pAry2(LneCnt, 19) = IIf(arrRecords2(2, PrtCnt) = "5L B", "-", arrRecords3(23, PrtCnt) & "")
'        pAry2(LneCnt, 20) = IIf(arrRecords3(24, PrtCnt) = 0, "-", arrRecords3(24, PrtCnt) & "") ' 0 JOMINY_RST2
'        pAry2(LneCnt, 21) = IIf(arrRecords3(25, PrtCnt) = 0, "-", arrRecords3(25, PrtCnt) & "") ' 0 JOMINY_RST3
        pAry2(LneCnt, 21) = IIf(arrRecords3(26, PrtCnt) = 0, "-", arrRecords3(26, PrtCnt) & "") ' GRAIN_SIZE_RST
        pAry2(LneCnt, 22) = arrRecords3(27, PrtCnt) & ""                ' NON_METAL_RST
        pAry2(LneCnt, 24) = arrRecords3(28, PrtCnt) & ""                ' BS_RST
        
        lSumQNTY = lSumQNTY + arrRecords2(4, PrtCnt)                    'PAGE SUM QUANTITY
        dSumWGT = dSumWGT + arrRecords2(5, PrtCnt)                      'PAGE SUM WEIGHT
        
        lSumQNTY_T = lSumQNTY_T + arrRecords2(4, PrtCnt)                    'TOTAL SUM QUANTITY
        dSumWGT_T = dSumWGT_T + arrRecords2(5, PrtCnt)                      'TOTAL SUM WEIGHT
        
        If LneCnt = 5 Then
            On Error GoTo ErrProc
            
            Set xlApp = GetObject("", "Excel.Application")
            If Err.Number = 429 Then
                Set xlApp = CreateObject("", "Excel.Application")
            End If
        
            xlApp.Workbooks.Open (App.Path & "\AQD060C.xls")
            Set xlSheet = xlApp.Worksheets("Sheet1")
            
            Call MillSheetPrint_P_Head(arrRecords1)
            
            xlSheet.Range("B30").Value = lSumQNTY   'arrRecords1(7, 0) & ""            'SUM_CNT
            xlSheet.Range("C30").Value = dSumWGT    'arrRecords1(8, 0) & ""            'SUM_WGT
'            xlSheet.Range("J10").Value = arrRecords3(25, PrtCnt)
                                   
            If arrRecords2(22, PrtCnt) = "Als" Then
               xlSheet.Range("V13").Value = "Als"
            End If
            
            If Trim(arrRecords1(4, 0)) = "06NGE/P020" Then
               xlSheet.Range("S13").Value = "Ca"
            End If
            
            If Trim(arrRecords1(4, 0)) = "06NGE/P084" Then
               xlSheet.Range("S13").Value = "Ca"
               xlSheet.Range("V13").Value = "Als"
            End If
            If PrtCnt = RowCNT Then
                xlSheet.Range("Q28").Value = "�ϼ�(Total)"
                xlSheet.Range("Q29").Value = "����  Piece"              'SUM_CNT UNIT
                xlSheet.Range("Q30").Value = lSumQNTY_T                 'SUM_CNT
                xlSheet.Range("R29").Value = "����(mt)   Weight"        'SUM_WGT UNIT
                xlSheet.Range("R30").Value = dSumWGT_T                  'SUM_WGT
            End If
            xlSheet.Range("K31").Value = Round((RowCNT + 3) / 5, 0) & " - " & Round((PrtCnt + 3) / 5, 0)
            Page_no = Round((PrtCnt + 3) / 5, 0)
            
            xlSheet.Range("B15:C19").Value = pAry11
            xlSheet.Range("D15:D19").Value = pAry12
            xlSheet.Range("F15:F19").Value = pAry13
            xlSheet.Range("G15:G19").Value = pAry14
            xlSheet.Range("J15:Y19").Value = pAry15
            xlSheet.Range("B23:Y27").Value = pAry2
            
            Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, Save_State, Save_Path)
            If Save_State = 0 Or Save_State = 1 Then
'            xlApp.Application.Visible = True
            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            End If

            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit

            LneCnt = 0
            dSumWGT = 0
            lSumQNTY = 0
            ReDim pAry11(1 To 5, 1 To 2)
            ReDim pAry12(1 To 5, 1 To 1)
            ReDim pAry13(1 To 5, 1 To 1)
            ReDim pAry14(1 To 5, 1 To 1)
            ReDim pAry15(1 To 5, 1 To 16)
            ReDim pAry2(1 To 5, 1 To 24)
            
        End If

    Loop Until PrtCnt = RowCNT
    
    If LneCnt <> 0 Then
    
        On Error GoTo ErrProc
        
        Set xlApp = GetObject("", "Excel.Application")
        If Err.Number = 429 Then
            Set xlApp = CreateObject("", "Excel.Application")
        End If
    
        xlApp.Workbooks.Open (App.Path & "\AQD060C.xls")
        Set xlSheet = xlApp.Worksheets("Sheet1")
            
        Call MillSheetPrint_P_Head(arrRecords1)

        xlSheet.Range("B30").Value = lSumQNTY   'arrRecords1(7, 0) & ""            'SUM_CNT
        xlSheet.Range("C30").Value = dSumWGT    'arrRecords1(8, 0) & ""            'SUM_WGT
        
        If arrRecords2(22, PrtCnt) = "Als" Then
           xlSheet.Range("V13").Value = "Als"
        End If
        
        If Trim(arrRecords1(4, 0)) = "06NGE/P020" Then
           xlSheet.Range("S13").Value = "Ca"
        End If
        
        If Trim(arrRecords1(4, 0)) = "06NGE/P084" Then
            xlSheet.Range("S13").Value = "Ca"
            xlSheet.Range("V13").Value = "Als"
        End If
        
        xlSheet.Range("Q28").Value = "�ϼ�(Total)"
        xlSheet.Range("Q29").Value = "����  Piece"              'SUM_CNT UNIT
        xlSheet.Range("Q30").Value = lSumQNTY_T                 'SUM_CNT
        xlSheet.Range("R29").Value = "����(mt)   Weight"        'SUM_WGT UNIT
        xlSheet.Range("R30").Value = dSumWGT_T                  'SUM_WGT
        xlSheet.Range("K31").Value = Round((RowCNT + 3) / 5, 0) & " - " & Round((PrtCnt + 3) / 5, 0)
        Page_no = Round((PrtCnt + 3) / 5, 0)

        
        xlSheet.Range("B15:C19").Value = pAry11
        xlSheet.Range("D15:D19").Value = pAry12
        xlSheet.Range("F15:F19").Value = pAry13
        xlSheet.Range("G15:G19").Value = pAry14
        xlSheet.Range("J15:Y19").Value = pAry15
        xlSheet.Range("B23:Y27").Value = pAry2
        
        Save_State = Cert_Save(xlApp.ActiveWorkbook, Cert_No, Page_no, Save_State, Save_Path)
        If Save_State = 0 Or Save_State = 1 Then
'           xlApp.Application.Visible = True
        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
        End If

        Set xlSheet = Nothing
        xlApp.ActiveWorkbook.Close False
        xlApp.Quit
        
    End If
            
    Set xlApp = Nothing
            
    Exit Function
    
ErrProc:
    If Err.Number = 429 Then
        MsgBox "Microsoft Excel Program Not Installed"
    Else
        MsgBox Err.Number & Err.Description
    End If
    MillSheetPrint_P = "ERROR"
    
    Set xlSheet = Nothing
    xlApp.ActiveWorkbook.Close False
    xlApp.Quit
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
    
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_P_Head
'   2.Name         : Pipe certificate print(Head table)
'   3.Input  Value : arrRecords1
'   4.Return Value :
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 06 .27
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.2
'--------------------------------------------------------------------------------------------------------
Private Sub MillSheetPrint_P_Head(ByVal arrRecords1 As Variant)
    Dim sDate           As String
    Dim sPONO           As String
    Dim sCert_No        As String
    Dim sPROD_CD        As String
    Dim sISP_SHP_NO     As String
    Dim sCONDITION_CN   As String
    Dim sCONDITION_EN   As String
    Dim sREMARK         As String
    
    sDate = GetPrintDate()
    sISP_SHP_NO = arrRecords1(5, 0)
    sCert_No = arrRecords1(0, 0)
    
    sPONO = GetPoNoLot(sISP_SHP_NO)
    sPROD_CD = Mid(sCert_No, 1, 2)
    sCONDITION_CN = GetConditionOfDelivery_C(sISP_SHP_NO)
    sCONDITION_EN = GetConditionOfDelivery_E(sISP_SHP_NO)
    sREMARK = GetRemark(sISP_SHP_NO)
    
    If sPONO = "N" Then
        If IsNull(arrRecords1(4, 0)) Then
            sPONO = ""
        Else
            sPONO = arrRecords1(4, 0) & ""
        End If
    Else
        sPONO = sPONO
    End If
    
    If sREMARK = "N" Then
        sREMARK = ""
    Else
        sREMARK = sREMARK
    End If
    
    If (sCONDITION_CN = "N" Or Len(Trim(sCONDITION_CN)) = 0) Or _
       (sCONDITION_EN = "N" Or Len(Trim(sCONDITION_EN)) = 0) Then
        sCONDITION_EN = "TMCP"
        sCONDITION_CN = "TMCP"
    End If
'--- PRINT ----------------------------------------------------------------------------------------------
    xlSheet.Range("C2").Value = arrRecords1(0, 0) & ""         'CERT_NO
    
    If UCase(sPROD_CD) = "HC" Then
        xlSheet.Range("B12").Value = "��Ʒ��" + Chr$(10) + "Coil ��"
        xlSheet.Range("B20").Value = "��Ʒ��" + Chr$(10) + "Coil ��"
    Else
        xlSheet.Range("B12").Value = "��Ʒ��" + Chr$(10) + "Plate ��"
        xlSheet.Range("B20").Value = "��Ʒ��" + Chr$(10) + "Plate ��"
    End If
    
    If Mid(sPONO, 1, 10) = "07JTE/P003" Then
       xlSheet.Range("C6").Value = "����Э��"
    End If
    
' SUN BIN START 20080522
    
    If Trim(arrRecords1(1, 0)) = "API SPEC 5L-2000" Then
        xlSheet.Range("C6").Value = "API SPEC 5L (43 ��)" & ""        'PROD_SPEC_NO
    Else
       If Mid(sPONO, 1, 19) = "WZG08-TJS06-XQ202NG" Or Mid(sPONO, 1, 19) = "WZG08-TJS06-XQ205NG" Or _
          Mid(sPONO, 1, 19) = "WZG08-TJS06-XQ203NG" Or Mid(sPONO, 1, 19) = "WZG08-TJS06-XQ204NG" Or _
          Mid(sPONO, 1, 19) = "WZG08-TJS06-XQ201NG" Then
          xlSheet.Range("C6").Value = "Q/SY GJX 0126-2007"
       Else
        xlSheet.Range("C6").Value = arrRecords1(1, 0) & ""         'PROD_SPEC_NO
       End If
    End If
' SUN BIN END 20080522
    
    xlSheet.Range("U8").Value = arrRecords1(2, 0) & ""         'CUST_NAME
    
    xlSheet.Range("U6").Value = sPONO & ""         'PONO
    
    
    xlSheet.Range("U4").Value = arrRecords1(5, 0) & ""         'TRNS_NO
    xlSheet.Range("N28").Value = sUserName                     'TEST_EMP
    xlSheet.Range("C8").Value = arrRecords1(9, 0) & ""         'PROD_NAME
    xlSheet.Range("C9").Value = arrRecords1(12, 0) & ""         'PROD_NAME_ENG
    xlSheet.Range("B10").Value = "����״̬�� " & sCONDITION_CN & "" 'COND_SUPPLY
    xlSheet.Range("B11").Value = "Condition of Supply�� " & sCONDITION_EN & "" 'COND_SUPPLY
    xlSheet.Range("J10").Value = arrRecords1(14, 0) & ""       'UST_NAME
    xlSheet.Range("J11").Value = arrRecords1(13, 0) & ""       'IMPACT_SIZE
    xlSheet.Range("F28").Value = arrRecords1(11, 0) & ""       'CAR NO
    xlSheet.Range("T10").Value = sDate                         'PRINT DATE
                
    
    
    If Trim(arrRecords1(0, 0)) = "PP200703100103" Or Trim(arrRecords1(0, 0)) = "PP200703070110" Or _
       Trim(arrRecords1(0, 0)) = "PP200703010005" Or Trim(arrRecords1(0, 0)) = "PP200703020120" Or _
       Mid(Trim(sPONO), 1, 10) = "06NGE/P110" Then
    
        xlSheet.Range("B31").Value = "According to EN 10204:2004 3.1"
    Else
        If Trim(sREMARK) <> "N" Then
            xlSheet.Range("B31").Value = sREMARK
        End If
    End If
End Sub
'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - funGetQuery_D
'   2.Name         : Slab BREAK-EVEN CARD
'   3.Input  Value : sHEAT_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 03 .21
'   7.Modify Date  : 2009.02.28 Sun Bin
'   8.Comment      : Public
'   9.Version      : 0.0.2
'--------------------------------------------------------------------------------------------------------
Public Function funGetQuery_D(ByVal sHEAT_NO As String) As String
        
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
    Dim arrRecords3 As Variant
    Dim AdoRs As adodb.Recordset
           
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT HEAT_NO,GF_STLGRD_DETAIL(STLGRD) AS STLGRD_NAME "
    sQuery = sQuery + ", SLAB_SIZE,ORD_LEN,QNTY,GF_EMPNAMEFIND(CHECKER_EMP) AS CHECKER_NAME"
    sQuery = sQuery + ", GF_EMPNAMEFIND(ANALYST_EMP) AS ANALYST_NAME, GF_EMPNAMEFIND(WEIGHTMAN_EMP) AS WEIGHTER_NAME"
    sQuery = sQuery + ", REMARK,TOTAL_WEIGHT,INS_DATE,INS_TIME"
    sQuery = sQuery + " FROM QP_SLAB_CARD_HEAD WHERE HEAT_NO = '" & sHEAT_NO & "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funGetQuery_D = "Err Database"
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = "    SELECT HEAT_NO,SLAB_NO ,CHECK_PASS,WGT,LEN "
    sQuery = sQuery + "  FROM QP_SLAB_CARD_DETAIL  WHERE HEAT_NO  = '" & sHEAT_NO & "' AND  PRINT_TIME LIKE TO_CHAR(SYSDATE,'YYYYMMDDHH24MI')||'%' "
    sQuery = sQuery + "  ORDER BY SLAB_NO"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funGetQuery_D = "Err Database"
        Exit Function
    End If
    arrRecords2 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = "    SELECT A.CHEM_COMP_CD,A.CHEM_RSLT ,B.CHEM_COMP_SEQ "
    sQuery = sQuery + "  FROM QP_CHEM_RSLT A ,QP_CHEM_SEQ B,QP_SLAB_CARD_CHARGE C,QP_NISCO_CHEM D"
    sQuery = sQuery + "  WHERE A.HEAT_NO  = '" & sHEAT_NO & "' AND A.CHEM_COMP_CD=B.CHEM_COMP_CD AND A.HEAT_NO=C.HEAT_NO AND C.STLGRD=D.STLGRD  "
    sQuery = sQuery + "  AND D.CHEM_COMP_CD=A.CHEM_COMP_CD ORDER BY B.CHEM_COMP_SEQ"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funGetQuery_D = "Err Database"
        Exit Function
    End If
    arrRecords3 = AdoRs.GetRows
    AdoRs.Close
       
    Set AdoRs = Nothing
    
    funGetQuery_D = MillSheetPrint_D(arrRecords1, arrRecords2, arrRecords3)
       
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - funGetQuery_D
'   2.Name         : Slab BREAK-EVEN CARD
'   3.Input  Value : sHEAT_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 03 .21
'   7.Modify Date  : 2009.02.28 Sun Bin
'   8.Comment      : Public
'   9.Version      : 0.0.2
'--------------------------------------------------------------------------------------------------------
Public Function funslabcardQuery(ByVal sHEAT_NO As String, ByVal vPRT_SEQ As Integer) As String
        
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
    Dim arrRecords3 As Variant
    Dim AdoRs As adodb.Recordset
           
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT A.HEAT_NO,GF_STLGRD_DETAIL(A.STLGRD) AS STLGRD_NAME "
    sQuery = sQuery + ", A.SLAB_SIZE, B.SLAB_LEN,A.SLAB_CNT,GF_EMPNAMEFIND(B.CHEM_EMP) AS CHECKER_NAME"
    sQuery = sQuery + ", GF_EMPNAMEFIND(A.PRT_EMP) AS ANALYST_NAME, GF_EMPNAMEFIND(A.PRT_EMP) AS WEIGHTER_NAME"
    sQuery = sQuery + ", A.REMARK,A.SLAB_WGT,A.PRT_DATE,A.PRT_TIME "
    sQuery = sQuery + " FROM QP_SLAB_CARD_PRT A,QP_SLAB_CARD_CHARGE B  "
    sQuery = sQuery + " WHERE A.HEAT_NO = '" & sHEAT_NO & "' AND  A.PRT_SEQ = '" & vPRT_SEQ & "' AND  B.HEAT_NO = A.HEAT_NO "
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funslabcardQuery = "Err Database"
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = "    SELECT HEAT_NO,SLAB_NO_SEQ ,'',WGT,LEN "
    sQuery = sQuery + "  FROM QP_SLAB_CARD_SLAB  WHERE HEAT_NO  = '" & sHEAT_NO & "' AND  PRT_SEQ = '" & vPRT_SEQ & "' "
    sQuery = sQuery + "  ORDER BY SLAB_NO_SEQ"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funslabcardQuery = "Err Database"
        Exit Function
    End If
    arrRecords2 = AdoRs.GetRows
    AdoRs.Close
    
    sQuery = "    SELECT A.CHEM_COMP_CD,A.CHEM_RSLT ,B.CHEM_COMP_SEQ "
    sQuery = sQuery + "  FROM QP_CHEM_RSLT A ,QP_CHEM_SEQ B,QP_SLAB_CARD_CHARGE C,QP_NISCO_CHEM D"
    sQuery = sQuery + "  WHERE A.HEAT_NO  = '" & sHEAT_NO & "' AND A.CHEM_COMP_CD=B.CHEM_COMP_CD AND A.HEAT_NO=C.HEAT_NO AND C.STLGRD=D.STLGRD  "
    sQuery = sQuery + "  AND D.CHEM_COMP_CD=A.CHEM_COMP_CD ORDER BY B.CHEM_COMP_SEQ"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        funslabcardQuery = "Err Database"
        Exit Function
    End If
    arrRecords3 = AdoRs.GetRows
    AdoRs.Close
       
    Set AdoRs = Nothing
    
    If MillSheetPrint_D(arrRecords1, arrRecords2, arrRecords3) = "" Then
        funslabcardQuery = ""
    Else
        funslabcardQuery = "Err Database"
    End If
       
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_D
'   2.Name         : Slab BREAK_EVEN CARD print(detail table)
'   3.Input  Value : arrRecords1 ,arrRecords2
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 03 .21
'   7.Modify Date  : 2009.02.28 Sun Bin
'   8.Comment      : Private
'   9.Version      : 0.0.2
'--------------------------------------------------------------------------------------------------------
Private Function MillSheetPrint_D(ByVal arrRecords1 As Variant, ByVal arrRecords2 As Variant, ByVal arrRecords3 As Variant) As String
    Dim RowCNT          As Long
    Dim ChemCnt         As Long
    Dim PrtCnt          As Long
    Dim LneCnt          As Long
    Dim pAry11()        As String                   'CHEM
    Dim pAry21()        As String                   'CHEM
    Dim pAry12()        As String                   'CHEM
    Dim pAry22()        As String                   'CHEM
    Dim pAry13()        As String                   '01#-14# SLAB'S SLAB_NO
    Dim pAry23()        As String                   '01#-14# SLAB'S WEIGHT
    Dim pAry14()        As String                   '01#-14# SLAB'S LENGTH
    Dim pAry15()        As String                   '15#-24# SLAB'S SLAB_NO
    Dim pAry25()        As String                   '01#-14# SLAB'S WEIGHT
    Dim pAry16()        As String                   '15#-24# SLAB'S LENGTH
    Dim i               As Integer
    

    If IsEmpty(arrRecords1) Or IsEmpty(arrRecords2) Or IsEmpty(arrRecords2) Then
       MillSheetPrint_D = "Err Data"
       Exit Function
    End If
    
    RowCNT = UBound(arrRecords2, 2)
    ChemCnt = UBound(arrRecords3, 2)
    
    
    
    PrtCnt = -1
    LneCnt = 0

    
    ReDim pAry11(1 To 1, 1 To 17)
    ReDim pAry21(1 To 1, 1 To 17)
    ReDim pAry12(1 To 1, 1 To 17)
    ReDim pAry22(1 To 1, 1 To 17)
    ReDim pAry13(1 To 1, 1 To 17)
    ReDim pAry23(1 To 1, 1 To 17)
    ReDim pAry14(1 To 1, 1 To 17)
    ReDim pAry15(1 To 1, 1 To 17)
    ReDim pAry25(1 To 1, 1 To 17)
    ReDim pAry16(1 To 1, 1 To 17)
    
    
    For i = 0 To ChemCnt
       If i <= 16 Then
            pAry11(1, i + 1) = arrRecords3(0, i) & ""
            pAry21(1, i + 1) = IIf(arrRecords3(1, i) < 1, "0" & arrRecords3(1, i), arrRecords3(1, i)) & ""
        Else
            pAry12(1, i - 16) = arrRecords3(0, i) & ""
            pAry22(1, i - 16) = IIf(arrRecords3(1, i) < 1, "0" & arrRecords3(1, i), arrRecords3(1, i)) & ""
        End If
    Next i
    
    For i = 0 To RowCNT
        If i <= 16 Then
            
            pAry13(1, i + 1) = arrRecords2(1, i) & ""
            pAry23(1, i + 1) = arrRecords2(3, i) & ""
            pAry14(1, i + 1) = arrRecords2(4, i) & ""
        Else
            pAry15(1, i - 16) = arrRecords2(1, i) & ""
            pAry25(1, i - 16) = arrRecords2(3, i) & ""
            pAry16(1, i - 16) = arrRecords2(4, i) & ""
        End If
    Next i
        
        Set xlApp = GetObject("", "Excel.Application")
        If Err.Number = 429 Then
            Set xlApp = CreateObject("", "Excel.Application")
        End If
    
        xlApp.Workbooks.Open (App.Path & "\AQD080C.xls")
        Set xlSheet = xlApp.Worksheets("Sheet1")
        
            Call MillSheetPrint_D_Head(arrRecords1)
                    
            
            xlSheet.Range("B4:R4").Value = pAry11
            xlSheet.Range("B5:R5").Value = pAry21
            xlSheet.Range("B6:R6").Value = pAry12
            xlSheet.Range("B7:R7").Value = pAry22
            xlSheet.Range("B8:R8").Value = pAry13
            xlSheet.Range("B9:R9").Value = pAry23
            xlSheet.Range("B10:R10").Value = pAry14
            xlSheet.Range("B11:R11").Value = pAry15
            xlSheet.Range("B12:R12").Value = pAry25
            xlSheet.Range("B13:R13").Value = pAry16
            
            
            xlSheet.Range("B24:R24").Value = pAry11
            xlSheet.Range("B25:R25").Value = pAry21
            xlSheet.Range("B26:R26").Value = pAry12
            xlSheet.Range("B27:R27").Value = pAry22
            xlSheet.Range("B28:R28").Value = pAry13
            xlSheet.Range("B29:R29").Value = pAry23
            xlSheet.Range("B30:R30").Value = pAry14
            xlSheet.Range("B31:R31").Value = pAry15
            xlSheet.Range("B32:R32").Value = pAry25
            xlSheet.Range("B33:R33").Value = pAry16
            
            xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Set xlSheet = Nothing
            xlApp.ActiveWorkbook.Close False
            xlApp.Quit
            
'    End If
        
    Set xlApp = Nothing
    
    Exit Function
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - MillSheetPrint_D_Head
'   2.Name         : Slab certificate print(Head table)
'   3.Input  Value : arrRecords1
'   4.Return Value :
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2006. 10 .11
'   7.Modify Date  : 2009.02.28 Sun Bin
'   8.Comment      : Private
'   9.Version      : 0.0.2
'--------------------------------------------------------------------------------------------------------
Private Sub MillSheetPrint_D_Head(arrRecords1 As Variant)
Dim s_DATE As String
Dim s_TIME As String

    s_DATE = Format(Trim(arrRecords1(10, 0)), "####-##-##")
    s_TIME = Format(Trim(arrRecords1(11, 0)), "##:##:##")
    
    xlSheet.Range("B3").Value = arrRecords1(0, 0) & ""         'HEAT_NO
    xlSheet.Range("E3").Value = arrRecords1(1, 0) & ""         'STLGRD_NAME
    xlSheet.Range("H3").Value = arrRecords1(2, 0) & ""         'SLAB_SIZE
    xlSheet.Range("K3").Value = arrRecords1(3, 0) & ""         'ORD_LEN
    xlSheet.Range("B14").Value = arrRecords1(8, 0) & ""         'REMARK
    xlSheet.Range("A16").Value = arrRecords1(4, 0) & ""         'ELIGIBLE_QNTY
    xlSheet.Range("Q16").Value = arrRecords1(5, 0) & ""         'CHECKER_EMP
    xlSheet.Range("G16").Value = arrRecords1(9, 0) & ""         'TOTAL_WEIGHT
'    xlSheet.Range("I16").Value = arrRecords1(7, 0) & ""         'WEIGHTMAN_EMP
'    xlSheet.Range("Q16").Value = arrRecords1(6, 0) & ""         'ANALYST_EMP
    xlSheet.Range("J18").Value = s_DATE & ""                    'INS_DATE
    xlSheet.Range("N18").Value = s_TIME & ""                    'INS_TIME
    
'-------------------------------------------------------------------------------
'*******************************************************************************
'-------------------------------------------------------------------------------
    xlSheet.Range("B23").Value = arrRecords1(0, 0) & ""         'HEAT_NO
    xlSheet.Range("E23").Value = arrRecords1(1, 0) & ""         'STLGRD_NAME
    xlSheet.Range("H23").Value = arrRecords1(2, 0) & ""         'SLAB_SIZE
    xlSheet.Range("K23").Value = arrRecords1(3, 0) & ""         'ORD_LEN
    xlSheet.Range("B34").Value = arrRecords1(8, 0) & ""         'REMARK
    xlSheet.Range("A36").Value = arrRecords1(4, 0) & ""         'ELIGIBLE_QNTY
    xlSheet.Range("Q36").Value = arrRecords1(5, 0) & ""         'CHECKER_EMP
    xlSheet.Range("G36").Value = arrRecords1(9, 0) & ""         'TOTAL_WEIGHT
'    xlSheet.Range("I36").Value = arrRecords1(7, 0) & ""         'WEIGHTMAN_EMP
'    xlSheet.Range("Q36").Value = arrRecords1(6, 0) & ""         'ANALYST_EMP
    xlSheet.Range("J38").Value = s_DATE & ""                    'INS_DATE
    xlSheet.Range("N38").Value = s_TIME & ""                    'INS_TIME
    
End Sub

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - GetConditionOfDelivery_C
'   2.Name         : Get Condition of delivery Chinese
'   3.Input  Value : sISP_SHP_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 06 .27
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function GetConditionOfDelivery_C(ByVal sISP_SHP_NO As String) As String
    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    sQuery = "Select COND_SUPPLY_CN From QP_CERT_LOT Where ISP_SHP_NO = " + "'" + sISP_SHP_NO + "'"
    
    Set AdoRs = New adodb.Recordset
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If AdoRs.EOF Then
        GetConditionOfDelivery_C = "N"
    Else
        If IsNull(AdoRs.Fields(0)) Then
            GetConditionOfDelivery_C = "N"
        Else
            GetConditionOfDelivery_C = AdoRs.Fields(0)
        End If
    End If
    
    AdoRs.Close
    
    Set AdoRs = Nothing
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - GetConditionOfDelivery_E
'   2.Name         : Get Condition of delivery Chinese
'   3.Input  Value : sISP_SHP_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 06 .27
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function GetConditionOfDelivery_E(ByVal sISP_SHP_NO As String) As String
    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    sQuery = "Select COND_SUPPLY_EN From QP_CERT_LOT Where ISP_SHP_NO = " + "'" + sISP_SHP_NO + "'"
    
    Set AdoRs = New adodb.Recordset
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If AdoRs.EOF Then
        GetConditionOfDelivery_E = "N"
    Else
        If IsNull(AdoRs.Fields(0)) Then
            GetConditionOfDelivery_E = "N"
        Else
            GetConditionOfDelivery_E = AdoRs.Fields(0)
        End If
    End If
    
    AdoRs.Close
    
    Set AdoRs = Nothing
End Function

'--------------------------------------------------------------------------------------------------------
'   1.ID           : basCertPrn - GetRemark
'   2.Name         : Get Condition of delivery Chinese
'   3.Input  Value : sISP_SHP_NO
'   4.Return Value : String
'   5.Writer       : Li Qing Yu
'   6.Create Date  : 2007. 06 .27
'   7.Modify Date  :
'   8.Comment      : Private
'   9.Version      : 0.0.1
'--------------------------------------------------------------------------------------------------------
Private Function GetRemark(ByVal sISP_SHP_NO As String) As String
    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    sQuery = "Select Remark From QP_CERT_LOT Where ISP_SHP_NO = " + "'" + sISP_SHP_NO + "'"
    
    Set AdoRs = New adodb.Recordset
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If AdoRs.EOF Then
        GetRemark = "N"
    Else
        If IsNull(AdoRs.Fields(0)) Then
            GetRemark = "N"
        Else
            GetRemark = AdoRs.Fields(0)
        End If
    End If
    
    AdoRs.Close
    
    Set AdoRs = Nothing
End Function


Private Function Old_Remark(ByVal sSTDSPEC As String, ByVal sPONO As String) As String
Dim sREMARK As String
    sREMARK = "N"
    
    If Mid(Trim(sPONO), 1, 10) = "07NGE/P007" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P111" Or _
       Mid(Trim(sPONO), 1, 10) = "06NGE/P110" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P069" Or _
       Mid(Trim(sPONO), 1, 10) = "06NGE/P073" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P029" Or _
       Mid(Trim(sPONO), 1, 10) = "06JTE/P028" Or Mid(Trim(sPONO), 1, 10) = "06JTE/P032" Or _
       Mid(Trim(sPONO), 1, 10) = "06JTE/P029" Or Mid(Trim(sPONO), 1, 10) = "06NGE/C004" Or _
       Mid(Trim(sPONO), 1, 10) = "06NGE/P107" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P104" Or _
       Mid(Trim(sPONO), 1, 10) = "06NGE/P102" Or Mid(Trim(sPONO), 1, 10) = "07NGE/P004" Or _
       Mid(Trim(sPONO), 1, 10) = "06JTE/P027" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P100" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P012" Or Mid(Trim(sPONO), 1, 12) = "06JTE/P004-1" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P016" Or Mid(Trim(sPONO), 1, 12) = "07JTE/P004-1" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P011" Or Mid(Trim(sPONO), 1, 10) = "07JTE/P010" Or _
       Mid(Trim(sPONO), 1, 10) = "07JTE/P012" Or Mid(Trim(sPONO), 1, 10) = "07NGE/P038" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P013" Or Mid(Trim(sPONO), 1, 10) = "07NGE/P029" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P046" Or Mid(Trim(sPONO), 1, 10) = "07JTE/P021" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P026" Or Mid(Trim(sPONO), 1, 10) = "07JTE/P015" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P037" Or Mid(Trim(sPONO), 1, 10) = "07JTE/P016" Or _
       Mid(Trim(sPONO), 1, 10) = "07NGE/P042" Or Mid(Trim(sPONO), 1, 10) = "07NGE/P030" Or _
       Mid(Trim(sPONO), 1, 10) = "06NGE/P040" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P095" Or _
       Mid(Trim(sPONO), 1, 10) = "06JTE/P022" Then
                    sREMARK = "According to EN 10204:2004 3.1"
    ElseIf Mid(Trim(sPONO), 1, 10) = "07JTE/P008" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P049" Or _
               Mid(Trim(sPONO), 1, 10) = "06NGE/P074" Then
                    sREMARK = "According to EN 10204 3.1B"
    ElseIf Mid(Trim(sPONO), 1, 10) = "06NGE/P081" Or Mid(Trim(sPONO), 1, 10) = "06NGE/P082" Then
                    sREMARK = "As per EN 10204 3.1.B"
    ElseIf Mid(Trim(sPONO), 1, 10) = "05NGE/P125" Then
                    sREMARK = "AS PER 3.1.B"
    ElseIf Mid(Trim(sPONO), 1, 10) = "06NGE/P075" And Mid(sSTDSPEC, 1, 2) = "EN" Then
            sREMARK = "According to EN 10204 3.1.C"
    ElseIf Mid(sSTDSPEC, 1, 2) = "EN" And (Mid(Trim(sPONO), 1, 10) = "06NGE/P086" Or _
                                            Mid(Trim(sPONO), 1, 10) = "06NGE/P087" Or _
                                            Mid(Trim(sPONO), 1, 10) = "06JTE/P029" Or _
                                            Mid(Trim(sPONO), 1, 10) = "07JTE/P005" Or _
                                            Mid(Trim(sPONO), 1, 10) = "06NGE/P102") Then
                    sREMARK = "According to EN 10204:2004 3.1"
    Else
                    sREMARK = "N"
    End If

    Old_Remark = sREMARK
End Function

'-------------------------------------------------------------------------------------------
' ������ʱ���
' ����
' 2014.9.3
'-------------------------------------------------------------------------------------------
Public Function funGetQuery_SPCL(ByVal sPacket_No As String, ByVal sSave_Path As String)
    Dim sQuery As String
    Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
    Dim AdoRs As adodb.Recordset
    Dim Save_State As String
    Dim i, II, j As Integer
    Dim RowCNT As Integer
    Dim sCert_No As String
    Dim RowFlag As Integer
    Dim Bend_str, COND_SUPPLY, STDSPEC, THK As String
    
    

    
    Set AdoRs = New adodb.Recordset
    sQuery = "{CALL AQD0012C.P_CERT_EXPORT('" + sPacket_No + "')}"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
     RowCNT = AdoRs.RecordCount
    AdoRs.Close
    Set AdoRs = Nothing
    
   
     
    Set xlApp = GetObject("", "Excel.Application")
    If Err.Number = 429 Then
      Set xlApp = CreateObject("", "Excel.Application")
    End If
    
    If arrRecords1(1, 0) = "EN 10025-6-S620QL1" Or arrRecords1(1, 0) = "JX/NG-0001-TB620" Or arrRecords1(1, 0) = "JX/NG-0002-TB620" Then
      xlApp.Workbooks.Open (App.Path & "\AQD055C.xls")
    Else
      xlApp.Workbooks.Open (App.Path & "\AQD054C.xls")
    End If
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    
    sCert_No = Format(Date, "YYMMDD") & "T" & sPacket_No 'Format(sPacket_No, "00000")
    
      STDSPEC = arrRecords1(1, 0)     '����״̬
      THK = Mid(arrRecords1(3, i), 1, InStr(arrRecords1(3, i), "*")) '���
      
      
      If InStr(STDSPEC, "S460") > 0 Or InStr(STDSPEC, "TB460") > 0 Then
        COND_SUPPLY = "����״̬: ����+���»ػ�"
      End If
      
      If InStr(STDSPEC, "S500") > 0 Or InStr(STDSPEC, "TB500") > 0 Then
      xlSheet.Range("J23").Value = "ReL��MPa"
        If Val(THK) > 16 Then
          COND_SUPPLY = "����״̬: ���+���»ػ�"
        Else
          COND_SUPPLY = "����״̬: TMCP+���»ػ�"
        End If
      End If
    
    
'---- EN10020-6-S620QL1 -----------------------------------------------------------------------------------
    If arrRecords1(1, 0) = "EN 10025-6-S620QL1" Or arrRecords1(1, 0) = "JX/NG-0001-TB620" Or arrRecords1(1, 0) = "JX/NG-0002-TB620" Then
    
      xlSheet.Range("O35").Value = arrRecords1(31, 0) '����ߴ�
      xlSheet.Range("C33").Value = RowCNT / 2 & "��"  '�ܼƿ���
      xlSheet.Range("AB8").Value = Date
      xlSheet.Range("AB6").Value = sCert_No
      
      Bend_str = "�����Ƕ�=" + Str(arrRecords1(51, 0)) + "��" + "������ֱ��=" + Str(arrRecords1(50, 0)) + "a���������ȡ�7a"
      xlSheet.Range("J34").Value = Bend_str
      
      
    
      RowFlag = 0
      
      j = 0                        '�ʱ���ɷ���
      II = i                       '�ʱ���������
      For i = 0 To RowCNT - 1      '��¼��

      '�ɷ�ֻȡ�����м�¼
    
        If i Mod 2 <> 0 Then
          xlSheet.Cells(j + 14, 1).Value = arrRecords1(2, i)    '��Ʒ��
          xlSheet.Cells(j + 14, 3).Value = arrRecords1(3, i)    '���
          xlSheet.Cells(j + 14, 7).Value = arrRecords1(4, i)    '�ߴ�
          xlSheet.Cells(j + 14, 8).Value = arrRecords1(32, i)   '����

          xlSheet.Cells(j + 14, 11).Value = IIf(Val(arrRecords1(6, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(6, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(6, i) & "") * 100, Val(arrRecords1(6, i) & "") * 100))           ' C_RST

          xlSheet.Cells(j + 14, 12).Value = IIf(Val(arrRecords1(7, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(7, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(7, i) & "") * 100, Val(arrRecords1(7, i) & "") * 100))           ' Si_RST

          xlSheet.Cells(j + 14, 13).Value = IIf(Val(arrRecords1(8, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(8, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(8, i) & "") * 100, Val(arrRecords1(8, i) & "") * 100))           ' MN_RST

          xlSheet.Cells(j + 14, 14).Value = IIf(Val(arrRecords1(9, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(9, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(9, i) & "") * 1000, Val(arrRecords1(9, i) & "") * 1000))           ' P_RST

          xlSheet.Cells(j + 14, 15).Value = IIf(Val(arrRecords1(10, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(10, i) & "") * 1000, 1, 1) = ".", _
          "0" & Val(arrRecords1(10, i) & "") * 1000, Val(arrRecords1(10, i) & "") * 1000))           ' S_RST

          xlSheet.Cells(j + 14, 16).Value = IIf(Val(arrRecords1(15, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(15, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(15, i) & "") * 100, Val(arrRecords1(15, i) & "") * 100))           ' NI_RST
          
          xlSheet.Cells(j + 14, 17).Value = IIf(Val(arrRecords1(14, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(14, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(14, i) & "") * 100, Val(arrRecords1(14, i) & "") * 100))           ' CR_RST

          xlSheet.Cells(j + 14, 18).Value = IIf(Val(arrRecords1(16, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(16, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(16, i) & "") * 100, Val(arrRecords1(16, i) & "") * 100))           ' MO_RST
          
          xlSheet.Cells(j + 14, 19).Value = IIf(Val(arrRecords1(12, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(12, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(12, i) & "") * 1000, Val(arrRecords1(12, i) & "") * 1000))           ' V_RST
        
          j = j + 1
        
        End If
      
        
        xlSheet.Cells(II + 21, 7).Value = arrRecords1(5, i)   '������
        xlSheet.Cells(II + 21, 8).Value = arrRecords1(20, i)   'ȡ��λ��
   
        xlSheet.Cells(II + 21, 9).Value = IIf(arrRecords1(21, i) = 0, "-", arrRecords1(21, i) & "")   ' YP_RST
        xlSheet.Cells(II + 21, 11).Value = IIf(arrRecords1(22, i) = 0, "-", arrRecords1(22, i) & "")  ' TS_RST
        xlSheet.Cells(II + 21, 13).Value = IIf(arrRecords1(23, i) = 0, "-", arrRecords1(23, i) & "")  ' EL_RST
        xlSheet.Cells(II + 21, 15).Value = IIf(arrRecords1(24, i) = 0, "-", arrRecords1(24, i) & "")  ' RA_RST
        

        xlSheet.Cells(II + 21, 17).Value = IIf(Val(arrRecords1(33, i) & "") = 0, "-", arrRecords1(33, i) & "")  ' Z������1
        xlSheet.Cells(II + 21, 18).Value = IIf(Val(arrRecords1(34, i) & "") = 0, "-", arrRecords1(34, i) & "")  ' Z������2
        xlSheet.Cells(II + 21, 19).Value = IIf(Val(arrRecords1(35, i) & "") = 0, "-", arrRecords1(35, i) & "")  ' Z������3
        xlSheet.Cells(II + 21, 20).Value = IIf(Val(arrRecords1(36, i) & "") = 0, "-", arrRecords1(36, i) & "")  ' Z��������ֵ

        xlSheet.Cells(II + 21, 21).Value = "�ϸ�"    '����

        xlSheet.Cells(II + 21, 22).Value = IIf(arrRecords1(26, i) = 0, "-", arrRecords1(26, i) & "")  ' IMPACT_TMP
        xlSheet.Cells(II + 21, 23).Value = IIf(arrRecords1(27, i) = 0, "-", arrRecords1(27, i) & "")  ' IMPACT_RST_1
        xlSheet.Cells(II + 21, 24).Value = IIf(arrRecords1(28, i) = 0, "-", arrRecords1(28, i) & "")  ' IMPACT_RST_2
        xlSheet.Cells(II + 21, 25).Value = IIf(arrRecords1(29, i) = 0, "-", arrRecords1(29, i) & "")  ' IMPACT_RST_3
        xlSheet.Cells(II + 21, 26).Value = IIf(arrRecords1(30, i) = 0, "-", arrRecords1(30, i) & "")  ' IMPACT_RST_AVE
        xlSheet.Cells(II + 21, 27).Value = IIf(arrRecords1(37, i) = 0, "-", arrRecords1(37, i) & "")  ' ��ά��1
        xlSheet.Cells(II + 21, 28).Value = IIf(arrRecords1(38, i) = 0, "-", arrRecords1(38, i) & "")  ' ��ά��2
        xlSheet.Cells(II + 21, 29).Value = IIf(arrRecords1(39, i) = 0, "-", arrRecords1(39, i) & "")  ' ��ά��3
        xlSheet.Cells(II + 21, 30).Value = IIf(arrRecords1(40, i) = 0, "-", arrRecords1(40, i) & "")  ' ��ά�ʾ�ֵ
        
        xlSheet.Cells(II + 22, 22).Value = IIf(arrRecords1(41, i) = 0, "-", arrRecords1(41, i) & "")  ' IMPACT_TMP �ڶ������������׷�ӳ����λ
        xlSheet.Cells(II + 22, 23).Value = IIf(arrRecords1(42, i) = 0, "-", arrRecords1(42, i) & "")  ' IMPACT_RST_1
        xlSheet.Cells(II + 22, 24).Value = IIf(arrRecords1(43, i) = 0, "-", arrRecords1(43, i) & "")  ' IMPACT_RST_2
        xlSheet.Cells(II + 22, 25).Value = IIf(arrRecords1(44, i) = 0, "-", arrRecords1(44, i) & "")  ' IMPACT_RST_3
        xlSheet.Cells(II + 22, 26).Value = IIf(arrRecords1(45, i) = 0, "-", arrRecords1(45, i) & "")  ' IMPACT_RST_AVE
        xlSheet.Cells(II + 22, 27).Value = IIf(arrRecords1(46, i) = 0, "-", arrRecords1(46, i) & "")  ' ��ά��1
        xlSheet.Cells(II + 22, 28).Value = IIf(arrRecords1(47, i) = 0, "-", arrRecords1(47, i) & "")  ' ��ά��2
        xlSheet.Cells(II + 22, 29).Value = IIf(arrRecords1(48, i) = 0, "-", arrRecords1(48, i) & "")  ' ��ά��3
        xlSheet.Cells(II + 22, 30).Value = IIf(arrRecords1(49, i) = 0, "-", arrRecords1(49, i) & "")  ' ��ά�ʾ�ֵ
        If arrRecords1(52, i) = "Y" Then
        xlSheet.Cells(II + 21, 31).Value = "�ϸ�"    '�Ͽ�
        End If
        xlSheet.Cells(II + 21, 32).Value = "�ϸ�"    '̽��
         II = II + 2

      Next i
      
'---- ������׼ --------------------------------------------------------------------------------------------------------------
    Else
    
      xlSheet.Range("P38").Value = arrRecords1(31, 0) '����ߴ�
      xlSheet.Range("C36").Value = RowCNT / 2 & "��"  '�ܼƿ���
      xlSheet.Range("Z8").Value = Date
      xlSheet.Range("Z6").Value = sCert_No
    
     xlSheet.Range("C38").Value = COND_SUPPLY
     Bend_str = "�����Ƕ�=" + Str(arrRecords1(51, 0)) + "��" + "������ֱ��=" + Str(arrRecords1(50, 0)) + "a���������ȡ�7a"
      xlSheet.Range("I37").Value = Bend_str
    
      j = 0
      For i = 0 To RowCNT - 1

      '�ɷ�ֻȡ�����м�¼
    
        If i Mod 2 <> 0 Then
          xlSheet.Cells(j + 14, 2).Value = arrRecords1(2, i)
          xlSheet.Cells(j + 14, 3).Value = arrRecords1(3, i)
          xlSheet.Cells(j + 14, 8).Value = arrRecords1(4, i)
          xlSheet.Cells(j + 14, 9).Value = arrRecords1(32, i)

          xlSheet.Cells(j + 14, 12).Value = IIf(Val(arrRecords1(6, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(6, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(6, i) & "") * 100, Val(arrRecords1(6, i) & "") * 100))           ' C_RST

          xlSheet.Cells(j + 14, 13).Value = IIf(Val(arrRecords1(7, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(7, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(7, i) & "") * 100, Val(arrRecords1(7, i) & "") * 100))           ' Si_RST

          xlSheet.Cells(j + 14, 14).Value = IIf(Val(arrRecords1(8, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(8, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(8, i) & "") * 100, Val(arrRecords1(8, i) & "") * 100))           ' MN_RST

          xlSheet.Cells(j + 14, 15).Value = IIf(Val(arrRecords1(9, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(9, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(9, i) & "") * 1000, Val(arrRecords1(9, i) & "") * 1000))           ' P_RST

          xlSheet.Cells(j + 14, 16).Value = IIf(Val(arrRecords1(10, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(10, i) & "") * 1000, 1, 1) = ".", _
          "0" & Val(arrRecords1(10, i) & "") * 1000, Val(arrRecords1(10, i) & "") * 1000))           ' S_RST
 
          xlSheet.Cells(j + 14, 17).Value = IIf(Val(arrRecords1(11, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(11, i) & "") * 1000, 1, 1) = ".", _
          "0" & Val(arrRecords1(11, i) & "") * 1000, Val(arrRecords1(11, i) & "") * 1000))           ' NB_RST

          xlSheet.Cells(j + 14, 18).Value = IIf(Val(arrRecords1(13, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(13, i) & "") * 1000, 1, 1) = ".", _
          "0" & Val(arrRecords1(13, i) & "") * 1000, Val(arrRecords1(13, i) & "") * 1000))           ' TI_RST

          xlSheet.Cells(j + 14, 19).Value = IIf(Val(arrRecords1(14, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(14, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(14, i) & "") * 100, Val(arrRecords1(14, i) & "") * 100))           ' CR_RST

          xlSheet.Cells(j + 14, 20).Value = IIf(Val(arrRecords1(15, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(15, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(15, i) & "") * 100, Val(arrRecords1(15, i) & "") * 100))           ' NI_RST

          xlSheet.Cells(j + 14, 21).Value = IIf(Val(arrRecords1(17, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(17, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(17, i) & "") * 100, Val(arrRecords1(17, i) & "") * 100))           ' CU_RST

          xlSheet.Cells(j + 14, 22).Value = IIf(Val(arrRecords1(18, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(18, i) & "") * 100, 1, 1) = ".", _
          "0" & Val(arrRecords1(18, i) & "") * 100, Val(arrRecords1(18, i) & "") * 100))           ' CEQ_RST

          xlSheet.Cells(j + 14, 23).Value = IIf(Val(arrRecords1(19, i) & "") = 0, _
          "-", IIf(Mid(Val(arrRecords1(19, i) & "") * 1000, 1, 1) = ".", _
          "0" & Val(arrRecords1(19, i) & "") * 1000, Val(arrRecords1(19, i) & "") * 1000))           ' AL_RST
          
          If InStr(STDSPEC, "S460") > 0 Or InStr(STDSPEC, "TB460") > 0 Then
          
            xlSheet.Cells(j + 14, 24).Value = IIf(Val(arrRecords1(12, i) & "") = 0, _
            "-", IIf(Mid(Val(arrRecords1(12, i) & "") * 100, 1, 1) = ".", _
            "0" & Val(arrRecords1(12, i) & "") * 1000, Val(arrRecords1(12, i) & "") * 1000))           ' V_RST
          
            xlSheet.Cells(j + 14, 25).Value = IIf(Val(arrRecords1(16, i) & "") = 0, _
            "-", IIf(Mid(Val(arrRecords1(16, i) & "") * 100, 1, 1) = ".", _
            "0" & Val(arrRecords1(16, i) & "") * 100, Val(arrRecords1(16, i) & "") * 100))           ' MO_RST
          
          Else
            xlSheet.Cells(j + 14, 24).Value = "-"
            xlSheet.Cells(j + 14, 25).Value = "-"
          End If
        
          j = j + 1
        
        End If
      
        xlSheet.Cells(i + 24, 8).Value = arrRecords1(5, i)
        xlSheet.Cells(i + 24, 9).Value = arrRecords1(20, i)


        xlSheet.Cells(i + 24, 10).Value = IIf(arrRecords1(21, i) = 0, "-", arrRecords1(21, i) & "")   ' YP_RST
        xlSheet.Cells(i + 24, 12).Value = IIf(arrRecords1(22, i) = 0, "-", arrRecords1(22, i) & "")  ' TS_RST
        xlSheet.Cells(i + 24, 14).Value = IIf(arrRecords1(23, i) = 0, "-", arrRecords1(23, i) & "")  ' EL_RST
        xlSheet.Cells(i + 24, 16).Value = IIf(arrRecords1(24, i) = 0, "-", arrRecords1(24, i) & "")  ' RA_RST
        xlSheet.Cells(i + 24, 18).Value = "�ϸ�"
        xlSheet.Cells(i + 24, 20).Value = IIf(arrRecords1(26, i) = 0, "-", arrRecords1(26, i) & "")  ' IMPACT_TMP
        xlSheet.Cells(i + 24, 22).Value = IIf(arrRecords1(27, i) = 0, "-", arrRecords1(27, i) & "")  ' IMPACT_RST_1
        xlSheet.Cells(i + 24, 24).Value = IIf(arrRecords1(28, i) = 0, "-", arrRecords1(28, i) & "")  ' IMPACT_RST_2
        xlSheet.Cells(i + 24, 26).Value = IIf(arrRecords1(29, i) = 0, "-", arrRecords1(29, i) & "")  ' IMPACT_RST_3
        xlSheet.Cells(i + 24, 28).Value = IIf(arrRecords1(30, i) = 0, "-", arrRecords1(30, i) & "")  ' IMPACT_RST_AVE
                If arrRecords1(52, i) = "Y" Then
        xlSheet.Cells(i + 24, 30).Value = "�ϸ�"   '�Ͽ�
        End If
    '    xlSheet.Cells(i + 24, 30).Value = "�ϸ�"

      Next i
    End If
'-------------------------------------------------------------------------------------------------------------------------



            
    Save_State = Cert_Save(xlApp.ActiveWorkbook, sCert_No, 1, 1, sSave_Path)
    Set xlSheet = Nothing
    xlApp.ActiveWorkbook.Close False
    xlApp.Quit
    Set xlApp = Nothing
            
            
    
End Function