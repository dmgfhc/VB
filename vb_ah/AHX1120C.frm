VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AHX1120C 
   Caption         =   "详细目的地选择_AHX1120C"
   ClientHeight    =   8280
   ClientLeft      =   3165
   ClientTop       =   3840
   ClientWidth     =   10185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10185
   Begin VB.TextBox txt_dest_cd 
      Height          =   285
      Left            =   990
      TabIndex        =   9
      Top             =   7740
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_form_nm 
      Height          =   285
      Left            =   6660
      TabIndex        =   8
      Top             =   7695
      Visible         =   0   'False
      Width           =   2040
   End
   Begin Threed.SSPanel pnl_result 
      Height          =   5865
      Left            =   0
      TabIndex        =   0
      Top             =   1665
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   10345
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPSpread.vaSpread ssResult 
         Height          =   5760
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   8655
         _Version        =   393216
         _ExtentX        =   15266
         _ExtentY        =   10160
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   -2147483633
         MaxCols         =   0
         MaxRows         =   0
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "AHX1120C.frx":0000
      End
   End
   Begin Threed.SSPanel pnl_condition 
      Height          =   1050
      Left            =   0
      TabIndex        =   2
      Top             =   585
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1852
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPSpread.vaSpread ssWhere 
         Height          =   1050
         Left            =   45
         TabIndex        =   3
         Top             =   -30
         Width           =   8655
         _Version        =   393216
         _ExtentX        =   15266
         _ExtentY        =   1852
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   -2147483633
         MaxCols         =   0
         MaxRows         =   0
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   2
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   1
         SpreadDesigner  =   "AHX1120C.frx":02D7
         UserResize      =   1
      End
   End
   Begin Threed.SSPanel pnl_button 
      Height          =   555
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   979
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmd_refresh 
         Height          =   420
         Left            =   45
         TabIndex        =   5
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "刷新"
      End
      Begin Threed.SSCommand cmd_selection 
         Height          =   420
         Left            =   1215
         TabIndex        =   6
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "选择"
      End
      Begin Threed.SSCommand cmd_close 
         Height          =   420
         Left            =   2385
         TabIndex        =   7
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   4210752
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "关闭"
      End
   End
End
Attribute VB_Name = "AHX1120C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Order Management System
'-- Sub_System Name
'-- Program Name
'-- Program ID        ABX1040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.06
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Private Sub cmd_close_Click()
        
    Unload Me
    
End Sub

Private Sub cmd_Selection_Click()

    If ssResult.ActiveRow > 0 Then
    
        Dim I As Integer
    
        For I = 0 To Forms.Count - 1
           If Forms(I).Name = txt_form_nm.Text Then
              ssResult.Row = ssResult.ActiveRow: ssResult.Col = 2
              Forms(I).txt_dest_detail.Text = ssResult.Text
              
           End If
        Next I
        
        Unload Me
        
    End If

End Sub

Private Sub Form_Activate()

   Dim sQuery As String
   Dim AdoRs As ADODB.Recordset
   Dim iRowCount As Long
   Dim iColcount As Long
   Dim ArrayRecords As Variant

   sQuery = "SELECT seq,dest_detail,cust_addr,post FROM bp_detail_dest "
   sQuery = sQuery + " where dest_cd='" & Trim(txt_dest_cd.Text) & "'"
   Debug.Print sQuery
   
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
     
     If AdoRs.BOF Or AdoRs.EOF Then
        ssResult.MaxRows = 0
        Call Gp_MsgBoxDisplay("There Is No Relevant Data", "I")
     Else
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 1) >= 0 Then
        
             ssResult.MaxRows = UBound(ArrayRecords, 2) + 1
        
            For iRowCount = 0 To ssResult.MaxRows - 1
            
                 ssResult.Row = iRowCount + 1
                
                For iColcount = 0 To ssResult.MaxCols - 1
                
                     ssResult.Col = iColcount + 1
                    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                         ssResult.Text = ""
                    Else
                         ssResult.Text = Trim(ArrayRecords(iColcount, iRowCount))
                    End If
                            
                Next iColcount
                
            Next iRowCount
            
         End If
        
         ssResult.ReDraw = True
         ssWhere.ReDraw = True
         Screen.MousePointer = vbDefault
      End If
 '颜色
    Call ssWhere_setting
    Call ssResult_setting

    Call Gp_Sp_ColGet(ssWhere, "B-System.INI", Me.Name, DD.DataDicType)
    Call Gp_Sp_ColGet(ssResult, "B-System.INI", Me.Name, DD.DataDicType)
    Call Gp_FormLoc_Get(Me, DD.DataDicType)


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 229 Then
        If KeyCode = 27 Then 'ESC Key
            Unload Me
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 229 Then
        If KeyCode = 27 Then 'ESC Key
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()


With ssWhere

 .MaxCols = 4
 .MaxRows = 1
 .Row = SpreadHeader
 .Col = 1
 .Row2 = SpreadHeader
 .Col2 = .MaxCols
 .Clip = "序号" + Chr(9) + "详细目的地" + Chr(9) + "地址" + Chr(9) + "邮编" + Chr(9)

End With

With ssResult

 .MaxCols = 4
 .Row = SpreadHeader
 .Col = 1
 .Row2 = SpreadHeader
 .Col2 = .MaxCols
' .Clip = "seq" + Chr(9) + "dest_detail" + Chr(9) + "cust_addr" + Chr(9) + "post" + Chr(9)
 .Clip = "序号" + Chr(9) + "详细目的地" + Chr(9) + "地址" + Chr(9) + "邮编" + Chr(9)
 
End With


End Sub

'================================='
' Where Spread Setting
'================================='
Public Sub ssWhere_setting()

    With ssWhere
    
        .RowHeight(-1) = 12.54
        .RowHeight(0) = 16
        
        .ColWidth(0) = 6
        .ColWidth(1) = 6
        .ColWidth(2) = 30
        .ColWidth(3) = 30
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .BlockMode = False
        
        .UserResize = UserResizeColumns
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        
        .SelBackColor = &H808040
        
        .Col = 0
        
        .Row = 0
        .Text = "项目"
        .Row = 1
        .Text = "值"
        
    End With
    
End Sub

'================================='
' Result Spread Setting
'================================='
Public Sub ssResult_setting()

    With ssResult
    
        .RowHeight(-1) = 12.54
        .RowHeight(0) = 18
        
        .ColWidth(0) = 6
        .ColWidth(1) = 6
        .ColWidth(2) = 30
        .ColWidth(3) = 30
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .BlockMode = False
        
        .UserResize = UserResizeColumns
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        
        .SelBackColor = &H808040
        
        .Col = 0
        .Col2 = -1
        .Row = 0
        .Row2 = -1
            
        .BlockMode = True
        .Lock = True
        .BlockMode = False
    
    End With
    
End Sub
'======================================'
' Form Resize --> Panel, Spread Resize
'======================================'
Private Sub Form_Resize()

    'Panel Width Setting
    If Me.ScaleWidth - pnl_button.Left > 1 Then
        pnl_button.Width = Me.ScaleWidth - 70
    End If
    
    If Me.ScaleWidth - pnl_condition.Left > 1 Then
        pnl_condition.Width = Me.ScaleWidth - 70
    End If
    
    If Me.ScaleWidth - pnl_result.Left > 1 Then
        pnl_result.Width = Me.ScaleWidth - 70
    End If

    'Panel Height Setting
    pnl_result.Height = Me.Height - 2250
    
    'Spread Width Setting (ssWhere)
    If Me.ScaleWidth - ssWhere.Left > 1 Then
        ssWhere.Width = Me.ScaleWidth - 180
    End If

    'Spread Width Setting (ssResult)
    If Me.ScaleWidth - ssResult.Left > 1 Then
        ssResult.Width = Me.ScaleWidth - 180
    End If
    
    'Spread Height Setting (ssResult)
    If Me.ScaleHeight - ssResult.Top > 1 Then
        ssResult.Height = pnl_result.Height - 100
    End If

End Sub




Private Sub ssResult_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(ssResult, Col, Row)

End Sub

Private Sub ssResult_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ssResult.MaxCols
    
        ssWhere.ColWidth(iCol) = ssResult.ColWidth(iCol)
        
    Next iCol

End Sub

Private Sub ssResult_DblClick(ByVal Col As Long, ByVal Row As Long)
 
   
    If Row > 0 Then
    
           Dim I As Integer
        
           For I = 0 To Forms.Count - 1
              If Forms(I).Name = txt_form_nm.Text Then
                  ssResult.Row = ssResult.ActiveRow: ssResult.Col = 2
                  Forms(I).txt_dest_detail.Text = ssResult.Text
                    
              End If
           Next I
         
        Unload Me
        
       
         
    End If
        
    
End Sub

Private Sub ssResult_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    
    If KeyCode = vbKeyReturn Then

        Dim I As Integer
    
        For I = 0 To Forms.Count - 1
           If Forms(I).Name = txt_form_nm.Text Then
                  ssResult.Row = ssResult.ActiveRow: ssResult.Col = 2
                  Forms(I).txt_dest_detail.Text = ssResult.Text
             
           End If
        Next I
            
        Unload Me
        
    End If

End Sub

Private Sub ssWhere_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)
    
    Dim iCol As Integer
    
    For iCol = 1 To ssWhere.MaxCols
    
        ssResult.ColWidth(iCol) = ssWhere.ColWidth(iCol)
        
    Next iCol
    
End Sub



