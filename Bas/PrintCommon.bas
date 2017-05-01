Attribute VB_Name = "PrintCommon"
Option Explicit

Global zoomIndex As Integer
Global prnSp     As vaSpread

Public Sub Gp_Sp_Print(prnTitle As String, prnSpread As vaSpread, Optional Arr_Header As Variant)
    Dim strHeader As String
    Dim iRow As Integer
    Dim iCol As Integer
    Dim i As Integer
    Dim j As Integer
    Dim lFontSize As Long
    
    Set prnSp = prnSpread
    
    prnSp.FontName = prnSpread.FontName
    prnSp.FontSize = prnSpread.FontSize
    
    lFontSize = 11
    With prnSp
        strHeader = "/fn" & Chr$(34) & .FontName & Chr$(34)
        strHeader = strHeader & "/fz" & Chr$(34) & 16 & Chr$(34) & "/fb1/fi1" & "/c" & prnTitle & _
                       "/n/n"
        If Not IsMissing(Arr_Header) Then
            iRow = UBound(Arr_Header, 1)
            iCol = UBound(Arr_Header, 2)
            
            If iRow > 9 Then iRow = 9
            If iCol > 2 Then iCol = 2
            
            For i = 0 To iRow
                
                For j = 0 To iCol
                    strHeader = strHeader & "/fn" & Chr$(34) & .FontName & Chr$(34)
                    strHeader = strHeader & "/fz" & Chr$(34) & lFontSize & Chr$(34) & _
                                "/fb0/fi0/fu0/fk0"
                    Select Case j
                    Case 0
                        strHeader = strHeader & "/l"
                    Case 2
                        strHeader = strHeader & "/c"
                    Case 4
                        strHeader = strHeader & "/r"
                    End Select
                    
                    strHeader = strHeader & Arr_Header(i, j)
                Next j
                
                strHeader = strHeader & "/n"
            Next i
            
        End If
        
        .PrintHeader = strHeader
        .PrintFooter = "/n/c/p" & "  of  " & .PrintPageCount & "/fz" & Chr$(34) & lFontSize & Chr$(34) & "/fb0/fi0" & "/r" & Now()
    
        If .EditMode = True Then
            .EditMode = False
            DoEvents
        End If
        
        .PrintMarginTop = 0.2 * 1440
        .PrintMarginBottom = 0.2 * 1440
        .PrintMarginLeft = 0.2 * 1440
        .PrintMarginRight = 0.2 * 1440
        .PrintColor = True
        
        'Init then zoom display
        zoomIndex = 8   'page height
    End With
    
    CommonPrint.Show vbModal
End Sub

Public Sub GetZoom(zoomlabel As Integer)
'Set up the print previews zoom
    
    With CommonPrint.prnPreview
        Select Case zoomlabel
            Case 0  '200%
                .PageViewType = 2
                .PageViewPercentage = 200
            Case 1  '150%
                .PageViewType = 2
                .PageViewPercentage = 150
            Case 2  '100%
                .PageViewType = 2
                .PageViewPercentage = 100
            Case 3  '75%
                .PageViewType = 2
                .PageViewPercentage = 75
            Case 4  '50%
                .PageViewType = 2
                .PageViewPercentage = 50
            Case 5  '25%
                .PageViewType = 2
                .PageViewPercentage = 25
            Case 6  '10%
                .PageViewType = 2
                .PageViewPercentage = 10
            Case 7  'Page Width
                .PageViewType = 3
            Case 8  'Page Height
                .PageViewType = 4
            Case 9  'Whole Page
                .PageViewType = 0
            Case 10 'Two Pages
                .PageViewType = 5
                .PageMultiCntH = 2
                .PageMultiCntV = 1
            Case 11 'Three Pages
                .PageViewType = 5
                .PageMultiCntH = 3
                .PageMultiCntV = 1
            Case 12 'Four Pages
                .PageViewType = 5
                .PageMultiCntH = 2
                .PageMultiCntV = 2
            Case 13 'Six Pages
                .PageViewType = 5
                .PageMultiCntH = 3
                .PageMultiCntV = 2
        End Select
   End With
End Sub


