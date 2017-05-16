Attribute VB_Name = "SpreadCommon"
Option Explicit

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Setting
'   2.Name         : Spread initialize Setting
'   3.Input  Value : sPname Variant, {MsgChk Boolean}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread initialize Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Setting(ByVal sPname As Variant, Optional MsgChk As Boolean = True)

    With sPname
    
        .RowHeight(-1) = 12.54
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .OperationMode = OperationModeRow
        .RetainSelBlock = True

        .UserResize = UserResizeColumns
        .AllowDragDrop = False
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
        
        'If .ColHeaderRows > 1 Then
        '    .Row = SpreadHeader + 1
        '    .FontBold = True
        'End If
        
        If MsgChk Then
            .LockBackColor = RGB(255, 255, 255)
        End If
        
        .MaxRows = 0
                
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ReadOnlySet
'   2.Name         : Spread Read Only Setting
'   3.Input  Value : sPname Variant
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Read Only Setting -- Locking
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ReadOnlySet(ByVal sPname As Variant)

    With sPname
            

        .Col = 0: .Col2 = .MaxCols
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .Lock = True
        .BlockMode = False
        .Protect = True
    
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ColGet
'   2.Name         : Spread ColWidth Read
'   3.Input  Value : sPname Variant, FileName String, sEcname String, {sType String}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread ColWidth Read -- .INI File
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ColGet(sPname As Variant, FileName As String, sEcname As String, Optional sType As String = "")

    Dim i As Integer
    Dim sKey As String
    Dim tPtext As String
    
    With sPname
    
        If .ColHeaderRows > 1 Then
        
            For i = 1 To .MaxCols
            
                .Col = i
            
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                tPtext = .Text
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                If .ColHidden = False Then
                    
                    If Trim(tPtext) <> Trim(.Text) Then
                        sKey = Trim(sType) + Trim(.Name) + "." + Trim(tPtext) + "(" + Trim(.Text) + ")"
                    Else
                        sKey = Trim(sType) + Trim(.Name) + "." + Trim(.Text)
                    End If
                    
                    .ColWidth(i) = GetPrivateProfileInt(sEcname, sKey, .ColWidth(i), App.Path & "\" & FileName)
                
                End If
            
            Next
            
        Else
        
            .Row = SpreadHeader + (.ColHeaderRows - 1)
                
            For i = 1 To .MaxCols
                .Col = i
                
                If .ColHidden = False Then
                    sKey = Trim(sType) + Trim(.Name) + "." + Trim(.Text)
                    .ColWidth(i) = GetPrivateProfileInt(sEcname, sKey, .ColWidth(i), App.Path & "\" & FileName)
                End If
            
            Next
            
        End If
        
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ColSet
'   2.Name         : Spread ColWidth Save
'   3.Input  Value : sPname Variant, FileName String, sEcname String, {sType String}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread ColWidth Save -- .INI File
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ColSet(sPname As Variant, FileName As String, sEcname As String, Optional sType As String = "")

    Dim i As Integer
    Dim sKey As String
    Dim sValue As String
    Dim tPtext As String

    With sPname
    
        If .ColHeaderRows > 1 Then
        
            For i = 1 To .MaxCols
                
                .Col = i
                
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                tPtext = .Text
                .Row = SpreadHeader + (.ColHeaderRows - 1)
            
                sValue = str$(.ColWidth(i))
                
                If .ColHidden = False Then
                    
                    If Trim(tPtext) <> Trim(.Text) Then
                        sKey = Trim(sType) + Trim(.Name) + "." + Trim(tPtext) + "(" + Trim(.Text) + ")"
                    Else
                        sKey = Trim(sType) + Trim(.Name) + "." + Trim(.Text)
                    End If
                    
                    Call WritePrivateProfileString(sEcname, sKey, sValue, App.Path & "\" & FileName)
                
                End If
            
            Next
            
        Else
        
            .Row = SpreadHeader + (.ColHeaderRows - 1)
            
            For i = 1 To .MaxCols
                .Col = i
                sValue = str$(.ColWidth(i))
                
                If .ColHidden = False Then
                    sKey = Trim(sType) + Trim(.Name) + "." + Trim(.Text)
                    Call WritePrivateProfileString(sEcname, sKey, sValue, App.Path & "\" & FileName)
                End If
            
            Next
            
        End If
    
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Spl_SizeGet
'   2.Name         : Splitter Height,Width Read
'   3.Input  Value : sPname Variant, FileName String, sEcname String, {sType String}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2010. 04 .14
'   7.Modify Date  :
'   8.Comment      : Splitter Height,Width Read -- .INI File
'---------------------------------------------------------------------------------------
Public Sub Gp_Spl_SizeGet(sPname As Variant, FileName As String, sEcname As String, Optional sType As String = "")

    Dim iCnt As Integer
    Dim sKeyH As String
    Dim sKeyW As String
    Dim sValueH As String
    Dim sValueW As String
    
    sKeyH = Trim(sPname.Name) + ".Height"
    sKeyW = Trim(sPname.Name) + ".Width"
    
    If sPname.Panes.Count = 1 Then Exit Sub
    
    For iCnt = 0 To sPname.Panes.Count - 1
        
        If sType = "H" Or sType = "" Then
            If sPname.Panes(iCnt).LockHeight = False Then
                sValueH = GetPrivateProfileInt(sEcname, sKeyH + "(" & iCnt & ")", 0, App.Path & "\" & FileName)
                If sValueH <> "0" Then
                    sPname.Panes(iCnt).Height = GetPrivateProfileInt(sEcname, sKeyH + "(" & iCnt & ")", 0, App.Path & "\" & FileName)
                End If
            End If
        End If
        
        If sType = "W" Or sType = "" Then
            If sPname.Panes(iCnt).LockWidth = False Then
                sValueW = GetPrivateProfileInt(sEcname, sKeyW + "(" & iCnt & ")", 0, App.Path & "\" & FileName)
                If sValueW <> "0" Then
                    sPname.Panes(iCnt).Width = GetPrivateProfileInt(sEcname, sKeyW + "(" & iCnt & ")", 0, App.Path & "\" & FileName)
                End If
            End If
        End If
        
    Next iCnt
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Spl_SizeSet
'   2.Name         : Splitter Height,Width Save
'   3.Input  Value : sPname Variant, FileName String, sEcname String, {sType String}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Splitter Height,Width Save -- .INI File
'---------------------------------------------------------------------------------------
Public Sub Gp_Spl_SizeSet(sPname As Variant, FileName As String, sEcname As String, Optional sType As String = "")

    Dim iCnt As Integer
    Dim sKeyH As String
    Dim sKeyW As String
    Dim sValueH As String
    Dim sValueW As String
    
    sKeyH = Trim(sPname.Name) + ".Height"
    sKeyW = Trim(sPname.Name) + ".Width"
    
    For iCnt = 0 To sPname.Panes.Count - 1
        
        sValueH = str$(sPname.Panes(iCnt).Height)
        sValueW = str$(sPname.Panes(iCnt).Width)
        
        Call WritePrivateProfileString(sEcname, sKeyH + "(" & iCnt & ")", sValueH, App.Path & "\" & FileName)
        Call WritePrivateProfileString(sEcname, sKeyW + "(" & iCnt & ")", sValueW, App.Path & "\" & FileName)
        
    Next iCnt
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ColHidden
'   2.Name         : Spread Column Hidden
'   3.Input  Value : sPname Variant, ColNum Variant, HiddenType Boolean
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Column Hidden
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ColHidden(sPname As Variant, ColNum As Variant, HiddenType As Boolean)

    With sPname
    
        .Col = ColNum
        .ColHidden = HiddenType
    
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_AutoInsert
'   2.Name         : Spread Row Auto Insert (Enter Key Press)
'   3.Input  Value : sPname Variant, {First_Col Variant}, {Last_Col Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Auto Insert (Enter Key Press)
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_AutoInsert(Sc As Collection)

    Dim iCount As Integer

    With Sc.Item("Spread")
    
        If .MaxRows < 1 Then Exit Sub
        
        If .ActiveCol <> IIf(Sc.Item("Last") = .MaxCols, .MaxCols, Sc.Item("Last")) Then Exit Sub
        If .ActiveRow <> .MaxRows Then Exit Sub
        
        .MaxRows = .MaxRows + 1
        
        .Row = .MaxRows
        .Action = SS_ACTION_INSERT_ROW
        .Col = 0: .Text = "Input"
        
        For iCount = 1 To .MaxCols
            .Col = iCount
            If .CellType = SS_CELL_TYPE_COMBOBOX Then .VALUE = 0
        Next iCount
        
        .Row = .MaxRows
        
        If Sc.Item("First") > 1 Then
            .Col = Sc.Item("First") - 1
        Else
            .Col = 0
        End If
        
        'Call Gp_Sp_ActiveCell(Sc.Item("Spread"), IIf(Sc.Item("first") > 1, Sc.Item("first") - 1, 0))
        
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_InAuthority
'   2.Name         : Spread Row Authority Insert
'   3.Input  Value : Sc Collection, Auth_Col Variant, {iType String}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .17
'   7.Modify Date  :
'   8.Comment      : Spread Row Authority Insert
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_InAuthority(Sc As Collection, Auth_Col As Variant, Optional iType As String)

    Dim iCount As Integer
    
    With Sc.Item("Spread")
    
        If iType = "" Then
            .Row = .ActiveRow
            .Col = 0
            
            If .Text = "Input" Or .Text = "Update" Then
                .Col = Auth_Col
                .Text = sUserID
                .Col = Auth_Col + 1
                .Text = sUserName
            End If
        Else
            For iCount = 1 To .MaxRows
                .Row = iCount
                .Col = Auth_Col
                .Text = sUserID
                .Col = Auth_Col + 1
                .Text = sUserName
            Next iCount
        End If
    
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_InsertRow
'   2.Name         : Spread Row Insert (Row Insert Key Press)
'   3.Input  Value : sPname Variant, {iRow Variant}, {RowHeader Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Insert (Row Insert Key Press)
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_InsertRow(sPname As Variant, Optional iRow As Variant = -1, Optional RowHeader As String = "Input")

    With sPname
    
        If iRow < 0 Then .Row = 1 Else .Row = iRow
        
        .ReDraw = False
        .Action = SS_ACTION_ACTIVE_CELL
        .EditMode = False: .MaxRows = .MaxRows + 1
        
        'Row = 0 is Default = 1
        If .MaxRows = 1 Then
            .Row = .ActiveRow
        Else
            .Row = .Row + 1
        End If
        
        .Action = SS_ACTION_INSERT_ROW
        .Col = 0: .Text = RowHeader
        .ReDraw = True
        
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_ProceExist
'   2.Name         : Spread Process Row Search
'   3.Input  Value : sPname Variant, {Tf Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Process Row Search
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_ProceExist(sPname As Variant, Optional Tf As Boolean = True) As Boolean

    Dim sMessg As String
    Dim lCount As Long
    Dim Proc As Long

    With sPname
    
        Proc = 0
        
        For lCount = 1 To .MaxRows
            .Col = 0: .Row = lCount
            If Trim(.Text) = "Input" Or Trim(.Text) = "Update" Or Trim(.Text) = "Delete" Then
                Proc = Proc + 1
                Exit For
            End If
        Next lCount
        
        If Proc > 0 Then
            If Tf Then
                sMessg = "表格中还有数据未处理，" + vbCrLf
                sMessg = sMessg + "放弃并继续吗？"
                
                If Gf_MessConfirm(sMessg, "Q") Then
                    Gf_Sp_ProceExist = False
                Else
                    Gf_Sp_ProceExist = True
                End If
                
            Else
                Gf_Sp_ProceExist = True
            End If
            
        Else
            Gf_Sp_ProceExist = False
        End If
    
        Exit Function
        
    End With

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_CellColor
'   2.Name         : Spread Cell Color Setting
'   3.Input  Value : sPname Variant, iCol Variant, iRow Variant, {fColor Variant},
'                    {bColor Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Cell Color Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_CellColor(sPname As Variant, iCol As Variant, iRow As Variant, Optional fColor As Variant = vbBlack, _
                           Optional bColor As Variant = vbWhite)

    With sPname

        .Col = iCol: .Col2 = iCol
        .Row = iRow: .Row2 = iRow
        
        .BlockMode = True
        .ForeColor = fColor
        .BackColor = bColor
        .BlockMode = False

    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ColColor
'   2.Name         : Spread Column Color Setting
'   3.Input  Value : sPname Variant, iCol Variant, {fColor Variant}, {bColor Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Column Color Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ColColor(sPname As Variant, iCol As Variant, Optional fColor As Variant = vbBlack, _
                          Optional bColor As Variant = vbWhite)

    With sPname
    
        .Col = iCol: .Col2 = iCol
        .Row = 1: .Row2 = -1
        
        .BlockMode = True
        .ForeColor = fColor
        .BackColor = bColor
        .BlockMode = False
        
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_HdColColor
'   2.Name         : Spread Column Header Color Setting
'   3.Input  Value : sPname Variant, iCol Variant
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 10 .29
'   7.Modify Date  :
'   8.Comment      : Spread Column Header Color Setting (F4 Function)
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_HdColColor(sPname As Variant, iCol As Variant)

    With sPname
    
        .Row = 0: .Row2 = 0
        .Col = iCol: .Col2 = iCol
        
        .BlockMode = True
        
        .CellType = SS_CELL_TYPE_STATIC_TEXT
        .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .TypeTextWordWrap = True
        
        .BackColor = &HE1E4CD
        .ForeColor = BLUE
        
        .BlockMode = False
        
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_RowColor
'   2.Name         : Spread Row Color Setting
'   3.Input  Value : sPname Variant, iRow Variant, {fColor Variant}, {bColor Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Color Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_RowColor(sPname As Variant, iRow As Variant, Optional fColor As Variant = vbBlack, _
                          Optional bColor As Variant = vbWhite)

    With sPname

        .Col = 1: .Col2 = -1
        .Row = iRow: .Row2 = iRow
        
        .BlockMode = True
        .ForeColor = fColor
        .BackColor = bColor
        .BlockMode = False

    End With

End Sub

'------------------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_BlockColor
'   2.Name         : Spread Block Color Setting
'   3.Input  Value : sPname Variant, iCol1 Variant, iCol2 Variant, iRow1 Variant, iRow2 Variant,
'                    {fColor Variant}, {bColor Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Block Color Setting
'------------------------------------------------------------------------------------------------
Public Sub Gp_Sp_BlockColor(sPname As Variant, iCol1 As Variant, iCol2 As Variant, iRow1 As Variant, _
                            iRow2 As Variant, Optional fColor As Variant = vbBlack, Optional bColor As Variant = vbWhite)

    With sPname
    
        .Col = iCol1: .Col2 = iCol2
        .Row = iRow1: .Row2 = iRow2
        
        .BlockMode = True
        .ForeColor = fColor
        .BackColor = bColor
        .BlockMode = False
        
    End With

End Sub

'-------------------------------------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_BlockLock
'   2.Name         : Spread Block Lock Setting
'   3.Input  Value : sPname Variant, iCol1 Variant, iCol2 Variant, iRow1 Variant, iRow2 Variant, {LockType Boolean}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Block Lock Setting
'--------------------------------------------------------------------------------------------------------------------
Public Sub Gp_Sp_BlockLock(sPname As Variant, iCol1 As Variant, iCol2 As Variant, iRow1 As Variant, iRow2 As Variant, _
                           Optional LockType As Boolean = True)

    With sPname

        If .Protect = False Then .Protect = True
        
        .Col = iCol1: .Col2 = iCol2
        .Row = iRow1: .Row2 = iRow2
        
        .BlockMode = True
        .Lock = LockType
        .BlockMode = False

    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ActiveCell
'   2.Name         : Spread Cel Active
'   3.Input  Value : sPname Variant, iCol Variant, {iRow Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Cel Active
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ActiveCell(sPname As Variant, Optional iCol As Variant = -1, Optional iRow As Variant = -1)

    With sPname
    
        If iCol > 0 Then .Col = iCol
        If iRow > 0 Then .Row = iRow
        
        .Action = SS_ACTION_ACTIVE_CELL
        .EditMode = True

    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Move
'   2.Name         : Control -> Spread Move aColumn
'   3.Input  Value : iCount Variant, Sc Collection, Mc Collection}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Cel Active
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Move(iCount As Variant, Sc As Collection, MC As Collection)

On Error GoTo Gp_Sp_Move_Error

    Dim i As Integer
    
    For i = 1 To MC.Item("aControl").Count
    
        If TypeOf MC.Item("aControl")(i) Is ComboBox Then  'ComboBox
        
            If MC.Item("aControl")(i).Style = 2 Then
            
                If MC.Item("aControl")(i).ListIndex = -1 Then
                    Call Gp_Sp_SendData(Sc.Item("Spread"), "0", Sc.Item("aColumn")(i), iCount)
                Else
                    Call Gp_Sp_SendData(Sc.Item("Spread"), MC.Item("aControl")(i).ListIndex, Sc.Item("aColumn")(i), iCount)
                End If
                
            Else
                Call Gp_Sp_SendData(Sc.Item("Spread"), MC.Item("aControl")(i).Text, Sc.Item("aColumn")(i), iCount)
            End If
            
        ElseIf TypeOf MC.Item("aControl")(i) Is CheckBox Then  'ComboBox
            Call Gp_Sp_SendData(Sc.Item("Spread"), MC.Item("aControl").Item(i), Sc.Item("aColumn").Item(i), iCount)
        ElseIf TypeOf MC.Item("aControl")(i) Is UDate Or TypeOf MC.Item("aControl")(i) Is sitxEdit Then  'Date
            Call Gp_Sp_SendData(Sc.Item("Spread"), MC.Item("aControl").Item(i).RawData, Sc.Item("aColumn").Item(i), iCount)
        Else
            Call Gp_Sp_SendData(Sc.Item("Spread"), MC.Item("aControl").Item(i).Text, Sc.Item("aColumn").Item(i), iCount)
        End If
        
    Next i
    
    Exit Sub
                    
Gp_Sp_Move_Error:

    Call Gp_MsgBoxDisplay("Gp_Sp_Move Error : " & Error)

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Ins
'   2.Name         : Spread Row Insert
'   3.Input  Value : Sc Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Insert
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Ins(Sc As Collection)

On Error GoTo Gp_Sp_Ins_Error

    Dim iCount As Integer
    
    Sc.Item("Spread").SetFocus
    Call Gp_Sp_InsertRow(Sc.Item("Spread"), Sc.Item("Spread").ActiveRow)
    Sc.Item("Spread").ReDraw = False
    
    For iCount = 1 To Sc.Item("iColumn").Count
    
        Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCount)
        If Sc.Item("Spread").CellType = SS_CELL_TYPE_COMBOBOX Then
            Sc.Item("Spread").VALUE = 0
        End If
        
    Next iCount
   
    Call Gp_Sp_ActiveCell(Sc.Item("Spread"), IIf(Sc.Item("first") > 0, Sc.Item("first"), 1))
    Sc.Item("Spread").ReDraw = True
    
    Exit Sub
    
Gp_Sp_Ins_Error:
    
    Call Gp_MsgBoxDisplay("Gp_Sp_Ins Error : " & Error)

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Cls
'   2.Name         : Spread Claer
'   3.Input  Value : Sc Collection
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Claer
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_Cls(Sc As Collection) As Boolean

On Error GoTo Gf_Sp_Cls_Error

    With Sc
        
        If Gf_Sp_ProceExist(.Item("Spread")) Then
            Gf_Sp_Cls = False
            Exit Function
        End If
        
        .Item("Spread").MaxRows = 0
        .Item("Spread").OperationMode = OperationModeNormal
        
        Gf_Sp_Cls = True
        
    End With

    Exit Function

Gf_Sp_Cls_Error:
    Gf_Sp_Cls = False

End Function

'-------------------------------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Collection
'   2.Name         : Spread Collection Setting
'   3.Input  Value : sPname Variant, Num Integer, pcol String, ncol String, mcol As String,
'                                                              iCol String, acol String, lCol String,
'                            pColumn Collection, nColumn Collection, mColumn Collection, iColumn Collection,
'                            aColumn Collection, lColumn Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Collection Setting
'--------------------------------------------------------------------------------------------------------------
Public Sub Gp_Sp_Collection(sPname As Variant, Num As Integer, pcol As String, ncol As String, mcol As String, _
                                                               iCol As String, acol As String, lCol As String, _
                            pColumn As Collection, nColumn As Collection, mColumn As Collection, iColumn As Collection, _
                            aColumn As Collection, lColumn As Collection, Optional iColumnColor As Boolean = True)
   
    If LCase(Trim(pcol)) = "p" Then       'PK Column
        pColumn.Add Item:=Num
    End If
    
    If LCase(Trim(ncol)) = "n" Then       'Necessary Column
        nColumn.Add Item:=Num
        'Call Gp_Sp_ColColor(SpName, Num, , &H80FF80)
    End If
    
    If LCase(Trim(mcol)) = "m" Then       'Spread Maxlength check Column
        mColumn.Add Item:=Num
    End If
    
    If LCase(Trim(iCol)) = "i" Then       'Spread Insert Column
        iColumn.Add Item:=Num
        'UPDATE BY KIM SUNG HO 07.08.17
        If iColumnColor Then Call Gp_Sp_ColColor(sPname, Num, , &HC0FFFF)
    End If
    
    If LCase(Trim(acol)) = "a" Then       'Master -> Spread Column
        aColumn.Add Item:=Num
        Call Gp_Sp_ColHidden(sPname, Num, True)
    End If
    
    If LCase(Trim(lCol)) = "l" Then       'Spread Lock Column
        lColumn.Add Item:=Num
        Call Gp_Sp_ColLock(sPname, Num, True)
    End If

    
End Sub

'-------------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Sort
'   2.Name         : Spread Column Sort
'   3.Input  Value : sPname Variant, Col Variant, Row Variant, {CL Boolean}, {KEY_COL Long}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Column Sort
'--------------------------------------------------------------------------------------------
Public Sub Gp_Sp_Sort(sPname As Variant, Col As Variant, Row As Variant, Optional CL As Boolean = False, Optional Key_Col As Long = 0)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sKey_Value() As String

    With sPname

        If .MaxRows < 1 Then Exit Sub
        
        If Row <= 0 And Col > 0 Then
        
            If CL And Key_Col <> 0 Then
            
                ReDim sKey_Value(1 To .MaxRows)
                        
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 0
                    
                    If .Text <> "" Then
                        j = j + 1
                        .Col = Key_Col
                        sKey_Value(j) = .Text
                        .Col = 0
                        .Text = ""
                        Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, i, i, BLACK, WHITE)
                    End If
                Next i
                
            Else
            
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 0
                    If .Text <> "" Then
                        Exit Sub
                    End If
                Next i
                
            End If
        
            .SortBy = SS_SORT_BY_ROW
            
            If .SortKey(1) = Col Then
                If .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING Then
                    .SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
                Else
                    .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
                End If
            Else
                If .SortKey(1) = -1 Then
                    .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
                End If
                .SortKey(1) = Col
                
            End If
            
            .Col = 1: .Col2 = .MaxCols
            .Row = 0: .Row2 = .MaxRows
            
            .Action = SS_ACTION_SORT
            
            'CLEAR
            If CL And Key_Col <> 0 Then
                For i = 1 To j
                    For k = 1 To .MaxRows
                        .Row = k
                        .Col = Key_Col
                        If .Text = sKey_Value(i) Then
                            Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, k, k, WHITE, BLUE)
                            .Col = 0
                            .Text = "Select"
                        End If
                    Next k
                Next i
            ElseIf CL And Key_Col = 0 Then
                .Col = 0: .Col2 = 0
                .Row = 1: .Row2 = .MaxRows
                .BlockMode = True
                .Text = ""
                .BlockMode = False
                Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, 1, .MaxRows, BLACK, WHITE)
            End If
            
        End If
        
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ColSort
'   2.Name         : Spread Column Multi Sort
'   3.Input  Value : sPname Variant, Sort_Form Form
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Column Multi Sort
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ColSort(sPname As Variant, Sort_Form As Form)

    Dim iCol As Integer
    Dim tPtext As String
    
    With sPname
    
        Sort_Form.cbo_first.AddItem ""
        Sort_Form.cbo_Second.AddItem ""
        Sort_Form.cbo_Third.AddItem ""
        
        For iCol = 1 To .MaxCols
        
            .Col = iCol
            '.Row = 0
            
            If .ColHeaderRows > 1 Then
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                tPtext = .Text
            
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                If .ColHidden = False Then
                    
                    If Trim(tPtext) <> Trim(.Text) Then
                        Sort_Form.cbo_first.AddItem Trim(tPtext) & "(" & .Text & ")" & Space(100) & iCol
                        Sort_Form.cbo_Second.AddItem Trim(tPtext) & "(" & .Text & ")" & Space(100) & iCol
                        Sort_Form.cbo_Third.AddItem Trim(tPtext) & "(" & .Text & ")" & Space(100) & iCol
                    Else
                        Sort_Form.cbo_first.AddItem .Text & Space(100) & iCol
                        Sort_Form.cbo_Second.AddItem .Text & Space(100) & iCol
                        Sort_Form.cbo_Third.AddItem .Text & Space(100) & iCol
                    End If
                End If
            
            Else
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                If .ColHidden = False Then
                    Sort_Form.cbo_first.AddItem .Text & Space(100) & iCol
                    Sort_Form.cbo_Second.AddItem .Text & Space(100) & iCol
                    Sort_Form.cbo_Third.AddItem .Text & Space(100) & iCol
                End If
            End If
            
        Next iCol
    
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ColLock
'   2.Name         : Spread Column Lock
'   3.Input  Value : sPname Variant, ColNum Variant, LockType Boolean
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Column Lock
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ColLock(sPname As Variant, ColNum As Variant, LockType As Boolean)

    With sPname
    
        .Protect = True
        .Col = ColNum: .Col2 = ColNum
        .Row = 1:      .Row2 = -1
        
        .BlockMode = True
        .Lock = LockType
        .BlockMode = False
    
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_CollectionLock
'   2.Name         : Spread Collection Column Lock
'   3.Input  Value : sPname Variant, lColumn Collection, {LockType Boolean}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Collection Column Lock
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_CollectionLock(sPname As Variant, lColumn As Collection, Optional LockType As Boolean = True)

    Dim iCount As Integer

    With sPname

        For iCount = 1 To lColumn.Count
        
            .Protect = True
            .Col = lColumn(iCount): .Col2 = lColumn(iCount)
            .Row = 1: .Row2 = .MaxRows
            
            .BlockMode = True
            .Lock = LockType
            .BlockMode = False
            
        Next iCount
    
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Change
'   2.Name         : Spread Change Check
'   3.Input  Value : Proc_Sc Collection, Sc Collection
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Change Check
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_Change(Proc_Sc As Collection, Sc As Collection) As Boolean
    
    Dim sMesg As String
 
    Proc_Sc("Sc").Item("Spread").Col = 0
    Proc_Sc("Sc").Item("Spread").Row = 0
    Proc_Sc("Sc").Item("Spread").Text = ""
    
    Sc("Spread").Col = 0
    Sc("Spread").Row = 0
    Sc("Spread").Text = "◎"

    If Proc_Sc("Sc").Item("Spread").Name = Sc.Item("Spread").Name Then Gf_Sp_Change = True: Exit Function
    
    'Process Row Check
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread"), False) Then
    
        sMesg = "请先处理当前表格中未处理的数据"
        
        Call Gp_MsgBoxDisplay(sMesg)
        
        Proc_Sc("Sc").Item("Spread").SetFocus
        Proc_Sc("Sc").Item("Spread").EditMode = True
        
        Sc("Spread").Col = 0
        Sc("Spread").Row = 0
        Sc("Spread").Text = "◎"
        
        Proc_Sc("Sc")("Spread").Col = 0
        Proc_Sc("Sc")("Spread").Row = 0
        Proc_Sc("Sc")("Spread").Text = ""
        
        Gf_Sp_Change = False
        Exit Function
        
    End If
    
    Proc_Sc.Remove ("Sc")
    Proc_Sc.Add Item:=Sc, Key:="Sc"
    
    Gf_Sp_Change = True
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ClipCopy
'   2.Name         : Spread Row Clipboard Copy
'   3.Input  Value : sPname Variant, {iRow Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Clipboard Copy
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ClipCopy(sPname As Variant, Optional iRow As Variant = -1)

    With sPname
    
        If .MaxRows < 1 Then
            Exit Sub
        Else
        
            .OperationMode = OperationModeNormal
            
            .Col = 1: .Col2 = .MaxCols
            
            If iRow = -1 Then
                .Row = 1: .Row2 = iRow
            ElseIf iRow < 0 Then
                .Row = .ActiveRow: .Row2 = .ActiveRow
            Else
                .Row = iRow: .Row2 = iRow
            End If
            
            .BlockMode = True
            .Action = SS_ACTION_SELECT_BLOCK
            .Action = SS_ACTION_CLIPBOARD_COPY
            .BlockMode = False
            
            .OperationMode = OperationModeRow
            
        End If
        
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_ClipPaste
'   2.Name         : Spread Row Clipboard Paste
'   3.Input  Value : sPname Variant, {iRow Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Clipboard Paste
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_ClipPaste(sPname As Variant, Optional iRow As Variant = -1)

    Dim sClipText As String

    With sPname
    
        .OperationMode = OperationModeNormal
        sClipText = Clipboard.GetText(vbCFText)
        
        If Trim(sClipText) <> "" Then
            .Col = 1: .Col2 = .MaxCols
            
            If iRow = -1 Then
                .Row = 1: .Row2 = iRow
            ElseIf iRow < 0 Then
                .Row = .ActiveRow: .Row2 = .ActiveRow
            Else
                .Row = iRow: .Row2 = iRow
            End If
            
            .BlockMode = True
            .Action = SS_ACTION_ACTIVE_CELL
            .Action = SS_ACTION_CLIPBOARD_PASTE
            .BlockMode = False
            
        End If
        
        .OperationMode = OperationModeRow
        
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Copy
'   2.Name         : Spread Row Copy
'   3.Input  Value : Sc Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Copy
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Copy(Sc As Collection)

    Call Gp_Sp_ClipCopy(Sc("Spread"), Sc("Spread").ActiveRow)

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Paste
'   2.Name         : Spread Row Paste
'   3.Input  Value : Sc Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Paste
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Paste(Sc As Collection)

    If Sc("Spread").MaxRows > 0 Then
        
        Call Gp_Sp_InsertRow(Sc("Spread"), Sc("Spread").ActiveRow)
        Call Gp_Sp_ClipPaste(Sc("Spread"), Sc("Spread").ActiveRow + 1)
        
        Sc.Item("Spread").RowHeight(Sc.Item("Spread").ActiveRow) = 12.54
        Call Gp_Sp_ActiveCell(Sc.Item("Spread"), IIf(Sc.Item("First") > 0, Sc.Item("First"), 1))
        
    End If

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Del
'   2.Name         : Spread Delete Marking
'   3.Input  Value : Sc Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Delete Marking
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Del(Sc As Collection)

    Dim i As Long
    
    With Sc.Item("Spread")
        
        If .MaxRows < 1 Then Exit Sub
        If .SelBlockRow < 1 Then Exit Sub
        
        For i = .SelBlockRow To .SelBlockRow2
            .Row = i
            .Col = 0
            
            If Trim(.Text) = "" Then
                .Text = "Delete"
            End If
        Next i
        
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_UpdateMake
'   2.Name         : Spread Update Marking
'   3.Input  Value : sPname Variant, Mode Integer
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Update Marking
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_UpdateMake(sPname As Variant, Mode As Integer)

    With sPname
    
        If .MaxRows < 1 Then Exit Sub
        
        .Col = .ActiveCol
        .Row = IIf(.ActiveRow > 0, .ActiveRow, 0)
        
        If Mode = 1 Then
            .Tag = .Text
        Else
            If Trim(.Tag) <> Trim(.Text) Then
                .Col = 0
                Select Case Trim(.Text)
                    Case "Input", "Update", "Delete"
                    Case Else
                        .Text = "Update"
                End Select
            End If
        End If
    
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_EventMake
'   2.Name         : Main Menu Click --> Spread Update Marking
'   3.Input  Value : sPname Variant
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Main Menu Click --> Spread Update Marking
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_EventMake(sPname As Variant)

    With sPname
    
        If Not (sPname Is Nothing) Then
        
            If .SelBlockRow = .SelBlockRow2 Then
                .Col = IIf(.ActiveCol > 0, .ActiveCol, 0)
                .Row = IIf(.ActiveRow > 0, .ActiveRow, 0)
                .Action = SS_ACTION_ACTIVE_CELL
            End If
        
        End If
    
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Cancel
'   2.Name         : Spread Row Cancel (Insert, Update, Delete)
'   3.Input  Value : Conn Connection, Sc Collection
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Cancel (Insert, Update, Delete)
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Cancel(Conn As ADODB.Connection, Sc As Collection)

On Error GoTo SpreadCancel_Error

    Dim sQuery As String
    Dim i As Integer
    Dim iRow, BR1, BR2 As Long

    With Sc
        
        Screen.MousePointer = vbHourglass
        .Item("Spread").ReDraw = False
        
        If .Item("Spread").MaxRows < 1 Or .Item("Spread").SelBlockRow < 1 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        BR1 = .Item("Spread").SelBlockRow
        BR2 = .Item("Spread").SelBlockRow2
        
        For iRow = .Item("Spread").SelBlockRow To BR2
            
            Select Case Trim(Gf_Sp_RcvData(.Item("Spread"), 0, iRow))
                
                Case "Input"
                    Call Gp_Sp_DeleteRow(.Item("Spread"), iRow)
                    iRow = iRow - 1: BR2 = BR2 - 1
                    
                    If iRow <> 0 Then
                        Call Gp_Sp_RowColor(Sc.Item("Spread"), iRow)
                        For i% = 1 To Sc!iColumn.Count
                            Call Gp_Sp_CellColor(.Item("Spread"), Sc!iColumn(i%), iRow, , &HC0FFFF)
                        Next i%
                    End If

                Case "Delete"
                    Call Gp_Sp_SendData(.Item("Spread"), "", 0, iRow)
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iRow)
                    
                    For i% = 1 To Sc!iColumn.Count
                        Call Gp_Sp_CellColor(.Item("Spread"), Sc!iColumn(i%), iRow, , &HC0FFFF)
                    Next i%
                    
                    For i% = 1 To Sc!lColumn.Count
                        Call Gp_Sp_CellColor(.Item("Spread"), Sc!lColumn(i%), iRow, , vbWhite)
                    Next i%
                    
                Case "Update"
                    sQuery = Gf_Sp_MakeQuery(.Item("Spread"), .Item("P-O"), "O", .Item("pColumn"), iRow)
                    Call Gp_Sp_OneRowDisplay(Conn, sQuery, .Item("Spread"), iRow)
                    Call Gp_Sp_SendData(.Item("Spread"), "", 0, iRow)
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iRow)
                    
                    For i% = 1 To Sc!iColumn.Count
                        Call Gp_Sp_CellColor(.Item("Spread"), Sc!iColumn(i%), iRow, , &HC0FFFF)
                    Next i%
                    
                    For i% = 1 To Sc!lColumn.Count
                        Call Gp_Sp_CellColor(.Item("Spread"), Sc!lColumn(i%), iRow, , vbWhite)
                    Next i%
                    
                Case Else
                    'sQuery = Gf_Sp_MakeQuery(.Item("Spread"), .Item("P-O"), "O", .Item("icolumn"), iRow)
                    'Call Gp_Sp_OneRowDisplay(Conn, sQuery, .Item("Spread"), iRow)
            End Select
            
            If iRow = BR2 Then
                Exit For
            End If

        Next iRow
        
        .Item("Spread").ReDraw = True
        
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
SpreadCancel_Error:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Gp_Sp_Cancel Error : " & Error)
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_RcvData
'   2.Name         : Spread Cell Recive  Data
'   3.Input  Value : sPname Variant, iCol Variant, iRow Variant
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Cell Recive  Data
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_RcvData(sPname As Variant, Optional iCol As Variant = -1, Optional iRow As Variant = -1) As Variant

    With sPname
    
        If iCol < 0 Then .Col = .Col Else .Col = iCol
        If iRow < 0 Then .Row = .Row Else .Row = iRow
        
        If .CellType = SS_CELL_TYPE_COMBOBOX Then
            Gf_Sp_RcvData = .VALUE
            
        ElseIf .CellType = SS_CELL_TYPE_CURRENCY Or .CellType = SS_CELL_TYPE_NUMBER Then
            Gf_Sp_RcvData = Val(.VALUE)
            
        Else
            Gf_Sp_RcvData = .Text
        End If
        
    End With
    
End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_SendData
'   2.Name         : Spread Cell Send  Data
'   3.Input  Value : sPname Variant, Indata Variant, {iCol Variant}, {iRow Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Cell Send  Data
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_SendData(sPname As Variant, Indata As Variant, Optional iCol As Variant = -1, Optional iRow As Variant = -1)

    With sPname

        If iCol = -1 Then .Col = .Col Else .Col = iCol
        If iRow < 0 Then .Row = .Row Else .Row = iRow
        
        If .CellType = SS_CELL_TYPE_COMBOBOX Or .CellType = SS_CELL_TYPE_CURRENCY Or .CellType = SS_CELL_TYPE_NUMBER Then
            .VALUE = IIf(VarType(Indata) = vbNull, 0, Indata)
        Else
            .SetText iCol, iRow, IIf(VarType(Indata) = vbNull, "", Indata)
        End If
    
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_DeleteRow
'   2.Name         : Spread Active Row Delete
'   3.Input  Value : sPname Variant, {iRow Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Active Row Delete
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_DeleteRow(sPname As Variant, Optional iRow As Variant = -1)

    With sPname

        If iRow < 0 Then .Row = .Row Else .Row = iRow
        
        .Action = SS_ACTION_DELETE_ROW
        .MaxRows = .MaxRows - 1

    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_MakeQuery
'   2.Name         : Spread Make Query
'   3.Input  Value : sPname Variant, ProcedureName Variant, iType String, QueryColumn Variant, iRow Variant
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_MakeQuery(sPname As Variant, ProcedureName As Variant, iType As String, QueryColumn As Variant, _
                                iRow As Variant) As String

On Error GoTo Sp_MakeQuery_Error
    
    Dim iCount As Integer
    Dim sQuery As String
    Dim sTemp As String
    Dim dTempFloat As Double
    Dim dTempInt As Double

    With sPname
    
        'Refer Or OneRow is No iType
        If iType = "R" Or iType = "O" Then
            sQuery = "{call " + ProcedureName + " ( "
        Else
            sQuery = "{call " + ProcedureName + " ( '" + iType + "',"
        End If
        
        .Row = iRow
        
        For iCount = 1 To QueryColumn.Count
        
            .Col = QueryColumn.Item(iCount)
            
            Select Case .CellType
            
                Case SS_CELL_TYPE_CURRENCY
                    If Trim(.Text) = "" Then
                        sQuery = sQuery + "0,"
                    Else
                        dTempFloat = .Text
                        sQuery = sQuery + str(dTempFloat) + ","
                    End If
                    
                Case SS_CELL_TYPE_NUMBER
                    If Trim(.Text) = "" Then
                        sQuery = sQuery + "0,"
                    Else
                        dTempInt = .Text
                        sQuery = sQuery + str(dTempInt) + ","
                    End If
                    
                Case SS_CELL_TYPE_CHECKBOX
                    If .Text = "1" Then
                        sQuery = sQuery + "'1',"
                    Else
                        sQuery = sQuery + "'0',"
                    End If
                    
                Case SS_CELL_TYPE_COMBOBOX
                    If Trim(.Text) = "" Then
                        sQuery = sQuery + "'0',"
                    Else
                        sQuery = sQuery + "'" + Trim(str(.VALUE)) + "',"
                    End If
                    
                Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                    If Trim(.Text) = "" Then
                        sQuery = sQuery + "'',"
                    Else
                        sQuery = sQuery + "'" + Trim(.VALUE) + "',"
                    End If
                    
                Case SS_CELL_TYPE_DATE
                    If Trim(.Text) = "" Then
                        sQuery = sQuery + "'',"
                    Else
                        sQuery = sQuery + "'" + Mid(Trim(.Text), 1, 4) & _
                                                Mid(Trim(.Text), 6, 2) & _
                                                Mid(Trim(.Text), 9, 2) + "',"
                    End If
                   
                Case Else
                    sTemp = Replace(.Text, "'", "''")
                    sQuery = sQuery + "'" + Trim(sTemp) + "',"
                    
            End Select
            
        Next iCount
    
        'Refer Or OneRow is Last String Delete
        If iType = "R" Or iType = "O" Then
            sQuery = Mid(sQuery, 1, Len(sQuery) - 1) + ")}"
        Else
            sQuery = sQuery + "?,?)}"
        End If

    End With

    Gf_Sp_MakeQuery = sQuery
    Exit Function

Sp_MakeQuery_Error:
    
    Gf_Sp_MakeQuery = "FAIL"
    Call Gp_MsgBoxDisplay("Gf_Sp_MakeQuery Error : " & Error)

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_OneRowDisplay
'   2.Name         : Spread One Row Display
'   3.Input  Value : Conn Connection, sQuery String, sPname Variant, {iRow Variant}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread One Row Display
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_OneRowDisplay(Conn As ADODB.Connection, sQuery As String, sPname As Variant, Optional iRow As Variant = -1)

On Error GoTo OneRowDisplay_Error

    Dim lCount As Long
    Dim iColcount As Integer
    Dim AdoRs As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Exit Sub
    End If
    
    With sPname

        If iRow = -1 Then lCount = .ActiveRow Else lCount = iRow
        
        .Row = iRow: .Col = 1: .Row2 = iRow: .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If Not AdoRs.BOF And Not AdoRs.EOF Then
        
            If Not AdoRs.EOF Then
                .Row = lCount
                
                For iColcount = 1 To .MaxCols
                    .Col = iColcount
                    
                    Select Case .CellType
                    
                        Case SS_CELL_TYPE_CHECKBOX
                            If VarType((AdoRs.Fields(iColcount - 1))) <> vbNull Or Trim(AdoRs.Fields(iColcount - 1)) = "1" Then
                                .Text = Trim(AdoRs.Fields(iColcount - 1))
                            End If
                            
                        Case SS_CELL_TYPE_COMBOBOX
                            If VarType((AdoRs.Fields(iColcount - 1))) = vbNull Or Trim(AdoRs.Fields(iColcount - 1)) = "" Then
                                .VALUE = 0
                            Else
                                .VALUE = Trim(AdoRs.Fields(iColcount - 1))
                            End If
                            
                        Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                            If VarType((AdoRs.Fields(iColcount - 1))) = vbNull Or Trim(AdoRs.Fields(iColcount - 1)) = "" Then
                                .VALUE = ""
                            Else
                                .VALUE = Trim(AdoRs.Fields(iColcount - 1))
                            End If
                            
                        Case SS_CELL_TYPE_DATE
                            If VarType((AdoRs.Fields(iColcount - 1))) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Mid(Trim(AdoRs.Fields(iColcount - 1)), 1, 4) & "-" & _
                                        Mid(Trim(AdoRs.Fields(iColcount - 1)), 5, 2) & "-" & _
                                        Mid(Trim(AdoRs.Fields(iColcount - 1)), 7, 2)
                            End If
                        
                        Case Else
                            If VarType((AdoRs.Fields(iColcount - 1))) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Trim(AdoRs.Fields(iColcount - 1))
                            End If
                            
                    End Select
                    
                Next iColcount
                
            End If
            
        End If

    End With

    AdoRs.Close
    Set AdoRs = Nothing
    Exit Sub
 
OneRowDisplay_Error:
    
    Set AdoRs = Nothing
    Call Gp_MsgBoxDisplay("Gp_Sp_OneRowDisplay Error : " & Error)
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Display
'   2.Name         : Spread Row Display
'   3.Input  Value : Conn Connection, sPname vaSpread, sQuery String, {lColumn Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Display
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_Display(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim sSpreadClip As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Gf_Sp_Display = True
        
        .ReDraw = False
        .MaxRows = 0: iCount = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Gf_Sp_Display = False
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
        
        '--- MODIFY 07.10.22 BY KIM SUNG HO
'        .Col = 1
'        .Col2 = .MaxCols
        '---------------------------------
    
        For iRowCount = 0 To .MaxRows - 1
        
            .Row = iRowCount + 1
            
            '--- MODIFY 07.10.22 BY KIM SUNG HO
'            .Row2 = iRowCount + 1
'            sSpreadClip = ""
            '----------------------------------
            
            For iColcount = 0 To .MaxCols - 1
            
                '--- MODIFY 07.10.22 BY KIM SUNG HO
'                sSpreadClip = sSpreadClip & ArrayRecords(iColcount, iRowCount) & Chr(9)
'
'            Next iColcount
'
'            '--- MODIFY 07.10.22 BY KIM SUNG HO
'            .Clip = sSpreadClip
                    
                .Col = iColcount + 1

                Select Case .CellType

                    Case SS_CELL_TYPE_CHECKBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                    Case SS_CELL_TYPE_COMBOBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                            .VALUE = 0
                        Else
                            .VALUE = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                    Case SS_CELL_TYPE_DATE
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                        End If

                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .VALUE = ""
                        Else
                            .VALUE = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                    Case Else
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                End Select

            Next iColcount
            
        Next iRowCount
            
        If Not lColumn Is Nothing Then

            'lControl Lock
            For iCount = 1 To lColumn.Count

                .Protect = True
                .Col = lColumn(iCount): .Col2 = lColumn(iCount)
                .Row = 1:               .Row2 = .MaxRows
                .BlockMode = True: .Lock = True
                .BlockMode = False

            Next iCount

        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Gf_Sp_Display = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Display Error : " & sQuery)
    Screen.MousePointer = vbDefault

End Function

'-----------------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Refer
'   2.Name         : Spread Refer
'   3.Input  Value : Conn Connection, Sc Collection, Mc Collection, {nCheckControl Collection},
'                                        {mCheckControl Collection},{MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Refer
'-----------------------------------------------------------------------------------------------
Public Function Gf_Sp_Refer(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                            Optional nCheckControl As Collection, Optional mCheckControl As Collection, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadRef_Error

    Dim sQuery As String
    Dim sMsg As String

'    If MsgChk Then
'        If Gf_Sp_ProceExist(Sc.Item("Spread")) Then
'            Gf_Sp_Refer = True
'            Exit Function
'        End If
'    End If

    If Not MC Is Nothing Then
    
        If Not nCheckControl Is Nothing Then
            sMsg = Gf_Ms_NeceCheck(nCheckControl)
            If sMsg <> "OK" Then
                sMsg = sMsg + "必须输入"
                Call Gp_MsgBoxDisplay(sMsg, "", "错误提示")
                Gf_Sp_Refer = False
                Exit Function
            End If
        End If
        
        If Not mCheckControl Is Nothing Then
            sMsg = Gf_Ms_NeceCheck2(mCheckControl)
            If sMsg <> "OK" Then
                sMsg = sMsg + "长度不正确"
                Call Gp_MsgBoxDisplay(sMsg, "", "错误提示")
                Gf_Sp_Refer = False
                Exit Function
            End If
        End If
        
    End If

    Sc.Item("Spread").OperationMode = OperationModeNormal
    
    If Not MC Is Nothing Then
        Gf_Sp_Refer = Gf_Sp_Display(Conn, Sc.Item("Spread"), Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", MC("pControl")), _
                                    Sc.Item("pColumn"), MsgChk)
        If Gf_Sp_Refer Then
           Call Gp_Ms_ControlLock(MC!lControl, True)
           MDIMain.StatusBar1.Panels(1) = "提示信息：查询成功"
        End If
    Else
        Gf_Sp_Refer = Gf_Sp_Display(Conn, Sc.Item("Spread"), Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), _
                                    "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), MsgChk)
    End If
    
    If Gf_Sp_Refer Then
        Sc.Item("Spread").OperationMode = OperationModeRow
        MDIMain.StatusBar1.Panels(1) = "提示信息：查询成功"
        'Sc!Spread.SetFocus
    End If
        
    Exit Function
    
SpreadRef_Error:

    Call Gp_MsgBoxDisplay("Gf_Sp_Refer Error : " & Error)
    Gf_Sp_Refer = False

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Process
'   2.Name         : Spread Data Process
'   3.Input  Value : Conn Connection, Sc Collection, Mc Collection, {RefChek Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Data Process
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

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

    Gf_Sp_Process = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
        Gf_Sp_Process = False
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
                    Gf_Sp_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Sp_Process = False: Exit Function
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
                adoCmd.Parameters(0).VALUE = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).VALUE = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).VALUE = "D"
                If Msg_Count = 1 Then
                   DelYN = Gf_MessConfirm("您确定要删除状态为[Delete]的数据吗？", "Q")
                   If DelYN Then Msg_Yes = "yes"
                   Msg_Count = Msg_Count + 1
                End If
                If Msg_Yes = "yes" Then DelYN = True
        End Select
          
        If ProcessChk = "YES" Or DelYN Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).VALUE = Trim(str(dTempFloat))
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).VALUE = Trim(str(dTempInt))
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").VALUE = "1" Then
                            adoCmd.Parameters(iCol).VALUE = "1"
                        Else
                            adoCmd.Parameters(iCol).VALUE = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = "0"
                        Else
                            adoCmd.Parameters(iCol).VALUE = Trim(str(Sc.Item("Spread").VALUE))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").VALUE) = "" Then
                            adoCmd.Parameters(iCol).VALUE = ""
                        Else
                            adoCmd.Parameters(iCol).VALUE = Trim(Sc.Item("Spread").VALUE)
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = ""
                        Else
                            adoCmd.Parameters(iCol).VALUE = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).VALUE = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Sp_Process = False
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
            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
                                                    Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", MC("pControl")), Sc.Item("pColumn"), False)
                                                    
        Else
            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
                           Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), False)
        End If
        
        MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了" & iProcessCount & "条记录"
        'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
        
    End If
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Gf_Sp_Process = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_Process = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_DelProcess
'   2.Name         : Header-Spread Data Delete Process
'   3.Input  Value : Conn Connection, Sc Collection
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Header-Spread Data Delete Process
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_DelProcess(Conn As ADODB.Connection, Sc As Collection) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim ProcessChk As String
    Dim dTempFloat As Double
    Dim dTempInt As Double
    Dim sTemp As String
    Dim adoCmd As ADODB.Command

    Gf_Sp_DelProcess = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function
    If Sc.Item("Spread").MaxRows < 1 Then
        Gf_Sp_DelProcess = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Sp_DelProcess = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input)
    For iCount = 0 To Sc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    For iCount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input"
                Call Gp_Sp_DeleteRow(Sc("Spread"), iCount)
                iCount = iCount - 1
                
            Case Else
                adoCmd.Parameters(0).VALUE = "D"
                ProcessChk = "YES"
            
        End Select
          
        If ProcessChk = "YES" Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).VALUE = str(dTempFloat)
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).VALUE = str(dTempInt)
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Text = "1" Then
                            adoCmd.Parameters(iCol).VALUE = "1"
                        Else
                            adoCmd.Parameters(iCol).VALUE = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = "0"
                        Else
                            adoCmd.Parameters(iCol).VALUE = Trim(str(Sc.Item("Spread").VALUE))
                        End If
                    
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = ""
                        Else
                            adoCmd.Parameters(iCol).VALUE = Trim(Sc.Item("Spread").VALUE)
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = ""
                        Else
                            adoCmd.Parameters(iCol).VALUE = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).VALUE = Trim(sTemp)
                        
                End Select
                
            Next iCol
            
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                Conn.RollbackTrans
                Gf_Sp_DelProcess = False
                Exit Function
        
             End If
            
        End If
        
    Next iCount
    
    Conn.CommitTrans
        
    Sc.Item("Spread").ReDraw = True
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_DelProcess = False
    Call Gp_MsgBoxDisplay("Gf_Sp_DelProcess Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_NeceCheck
'   2.Name         : Spread  Necessary Check
'   3.Input  Value : sPname Variant, iRow Variant, iCol Collection
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread  Necessary Check
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_NeceCheck(sPname As Variant, iRow As Variant, iCol As Collection) As String

    Dim iCount As Integer

    With sPname

        .Row = iRow
        
        For iCount = 1 To iCol.Count
        
            .Col = iCol.Item(iCount)
            
            Select Case .CellType
            
                Case SS_CELL_TYPE_COMBOBOX
                    If .TypeComboBoxEditable = True Then
                        If Trim(.Text) = "" Then
                            .Row = 0: Gf_Sp_NeceCheck = .Text: Exit Function
                        End If
                    Else
                        If Trim(.Text) = "" Or .VALUE = "0" Then
                            .Row = 0: Gf_Sp_NeceCheck = .Text: Exit Function
                        End If
                    End If
                    
                Case SS_CELL_TYPE_EDIT, SS_CELL_TYPE_DATE
                    If Len(Trim(.Text)) < .TypeEditLen Then
                        .Row = 0: Gf_Sp_NeceCheck = .Text: Exit Function
                    End If
                    
                Case Else
                    If Trim(.Text) = "" Then
                        If Trim(.Text) = "" Then
                            .Row = 0: Gf_Sp_NeceCheck = .Text: Exit Function
                        End If
                    End If
                    
            End Select
            
        Next iCount
        
        Gf_Sp_NeceCheck = "OK"

    End With

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_NeceCheck2
'   2.Name         : Spread  Necessary, MaxLength Check
'   3.Input  Value : sPname Variant, mColumn Collection, iRow Variant, iCol Collection
'   4.Return Value : String
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread  Necessary, MaxLength Check
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_NeceCheck2(ByVal sPname As Variant, mColumn As Collection, iRow As Variant, iCol As Collection) As String

    Dim iCount As Integer
    Dim iLeng As Integer

    With sPname

        .Row = iRow
        
        For iCount = 1 To iCol.Count
        
            .Col = iCol.Item(iCount)
            
            Select Case .CellType
            
                Case SS_CELL_TYPE_COMBOBOX
                    If .TypeComboBoxEditable = True Then
                        If Trim(.Text) = "" Then
                            .Row = 0: Gf_Sp_NeceCheck2 = .Text: Exit Function
                        End If
                    Else
                        If Trim(.Text) = "" Or .VALUE = "0" Then
                            .Row = 0: Gf_Sp_NeceCheck2 = .Text: Exit Function
                        End If
                    End If
                    
                Case SS_CELL_TYPE_EDIT, SS_CELL_TYPE_DATE
                    If Trim(.Text) = "" Then
                        .Row = 0: Gf_Sp_NeceCheck2 = .Text: Exit Function
                    End If
                    
                Case Else
                    If Trim(.Text) = "" Then
                        If Trim(.Text) = "" Then
                            .Row = 0: Gf_Sp_NeceCheck2 = .Text: Exit Function
                        End If
                    End If
                    
            End Select
            
        Next iCount
        
        'MAXLENGTH Check
        For iCount = 1 To mColumn.Count

            .Col = mColumn.Item(iCount)

            If .CellType <> SS_CELL_TYPE_COMBOBOX Then

                iLeng = .TypeMaxEditLen

                If Trim(.Text) <> "" And Len(Trim(.Text)) < iLeng Then
                    .Row = 0
                    Gf_Sp_NeceCheck2 = "FALSE" + "'" + .Text + "'"
                    Exit Function
                End If

            End If

        Next iCount
        
    End With

    Gf_Sp_NeceCheck2 = "OK"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Excel
'   2.Name         : Spread --> Excel
'   3.Input  Value : Fm Form, sPname Variant, bLkcol1 Long, bLkcol2 Long, bLkrow1 Long, bLkrow2 Long
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread --> Excel
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Excel(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .Row = bLkrow1: .Row2 = bLkrow2
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
                        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .Row = 1

            iExlCol = ColIndex
'            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1" Or Left(.Text, 1) = "7") And _
'               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
            If .CellType = SS_CELL_TYPE_EDIT Then
                If ColIndex > 104 Then
                    sExlRange1 = "D"
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C"
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B"
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If

                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14 Then
                     xlSheet.Range(sExlRange).NumberFormat = "@"
                End If
            End If
        Next
        
        xlSheet.Range("A1").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
            
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel", "W")

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_EvenRowBackcolor
'   2.Name         : Spread Even Odd Row Color Setting
'   3.Input  Value : sPname Variant, {MaxCnt Integer}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Even Odd Row Color Setting
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_EvenRowBackcolor(ByVal sPname As Variant, Optional MaxCnt As Integer = 0)
    
    Dim i As Integer
    
    With sPname
        .ReDraw = False
        
        For i = 1 To .MaxRows - MaxCnt
            .Row = i
            
            If i Mod 2 <> 0 Then
                .BlockMode = True
                .Row2 = i
                .Col = 1: .Col2 = -1
                .BackColor = &HF2F2F2   'RGB(241, 236, 255)   '&HFFC0FF
                .BlockMode = False
            Else
                .BlockMode = True
                .Row2 = i
                .Col = 1: .Col2 = -1
                .BackColor = &HFFFFFF
                .BlockMode = False
            End If
            
        Next i
        
        .ReDraw = True
        .Refresh
        
    End With
    
End Sub

'--------------------------------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_ColSum
'   2.Name         : Spread Column Sum
'   3.Input  Value : sPname Variant, iCol Long, {Start_Row Long}, {End_Row Long}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Column Sum
'--------------------------------------------------------------------------------------------------------------
Public Function Gf_Sp_ColSum(ByVal sPname As Variant, iCol As Long, Optional Start_Row As Long = 1, _
                                                                    Optional End_Row As Long = 0) As Double
        
    Dim lCount As Long
    Dim dSum As Double
    
    With sPname
    
        If End_Row > .MaxRows Or End_Row = 0 Then
            End_Row = .MaxRows
        End If
        
        .Col = iCol
        
        For lCount = Start_Row To End_Row
            .Row = lCount
            If .Text <> "" Then
                dSum = dSum + .VALUE
            End If
        Next lCount
    
    End With
    
    Gf_Sp_ColSum = dSum
    
End Function

'--------------------------------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_RowSum
'   2.Name         : Spread Row Sum
'   3.Input  Value : sPname Variant, iRow Long, {Start_Column Long}, {End_Column Long}
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Sum
'--------------------------------------------------------------------------------------------------------------
Public Function Gf_Sp_RowSum(ByVal sPname As Variant, iRow As Long, Optional Start_Col As Long = 1, _
                                                                    Optional End_Col As Long = 0) As Double
        
    Dim lCount As Long
    Dim dSum As Double
    
    With sPname
    
        If End_Col > .MaxCols Or End_Col = 0 Then
            End_Col = .MaxCols
        End If
        
        .Row = iRow
        
        For lCount = Start_Col To End_Col
            .Col = lCount
            If .Text <> "" Then
                dSum = dSum + .VALUE
            End If
        Next lCount
    
    End With
    
    Gf_Sp_RowSum = dSum
    
End Function
