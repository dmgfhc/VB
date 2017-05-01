VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form DataDic 
   Caption         =   "代码表"
   ClientHeight    =   7620
   ClientLeft      =   6075
   ClientTop       =   2010
   ClientWidth     =   5325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   5325
   Begin Threed.SSPanel pnl_result 
      Height          =   5865
      Left            =   45
      TabIndex        =   2
      Top             =   1710
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   10345
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPSpread.vaSpread ssResult 
         Height          =   1545
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   5145
         _Version        =   393216
         _ExtentX        =   9075
         _ExtentY        =   2725
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "DataDic.frx":0000
         UserResize      =   1
      End
   End
   Begin Threed.SSPanel pnl_condition 
      Height          =   1050
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   1852
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPSpread.vaSpread ssWhere 
         Height          =   960
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   5145
         _Version        =   393216
         _ExtentX        =   9075
         _ExtentY        =   1693
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         RowHeaderDisplay=   2
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   1
         SpreadDesigner  =   "DataDic.frx":0234
         UserResize      =   1
      End
   End
   Begin Threed.SSPanel pnl_button 
      Height          =   555
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   979
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmd_refresh 
         Height          =   420
         Left            =   45
         TabIndex        =   3
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         TabIndex        =   4
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         TabIndex        =   5
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   4210752
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
Attribute VB_Name = "DataDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Data Dictionary
'-- Sub_System Name
'-- Program Name
'-- Program ID        DataDic
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.06
'-- Modify            2003.7.31
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Dim PrevRow As Long      'Mouse previous Move Row

Private Sub cmd_close_Click()
        
    Unload Me
    
End Sub

Private Sub cmd_refresh_Click()

    Dim sQuery As String
    
    DD.DicRefType = "R"                 'DataDic Form Refer
    
    ssWhere.Row = 1
    
    Select Case DD.DataDicType
    
            Case "M"  'Common Code
            
                ssWhere.Col = 1
                sQuery = "          AND CD                        like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(CD_SHORT_NAME,'%')    like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 3
                sQuery = sQuery + " AND NVL(CD_NAME,'%')          like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 4
                sQuery = sQuery + " AND NVL(CD_SHORT_ENG,'%')     like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 5
                sQuery = sQuery + " AND NVL(CD_FULL_ENG,'%')      like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  CD  ASC "
                
            Case "U"  'Order Usage Code
            
                ssWhere.Col = 1
                sQuery = "          AND ENDUSE_CD                 like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(ENDUSE_NAME,'%')      like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  ENDUSE_CD  ASC "
            
            Case "C"  'Customer Code
            
                ssWhere.Col = 1
                sQuery = " WHERE    CUST_CD                       like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(CUST_NM,'%')          like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 3
                sQuery = sQuery + " AND NVL(CUST_NM_ENG,'%')      like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  CUST_CD  ASC "
                
            Case "D"  'Customer Destination Code
            
                ssWhere.Col = 1
                sQuery = " WHERE    DEST_CD                       like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(CITY_CD,'%')          like '" & Trim(ssWhere.Text) & "%' "
            
                ssWhere.Col = 3
                sQuery = sQuery + " AND NVL(STATION_CD,'%')       like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 4
                sQuery = sQuery + " AND NVL(DEST_NM,'%')          like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 5
                sQuery = sQuery + " AND NVL(DEST_NM_ENG,'%')      like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 6
                sQuery = sQuery + " AND NVL(DEST_ADDR,'%')        like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 7
                sQuery = sQuery + " AND NVL(POST,'%')             like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 8
                sQuery = sQuery + " AND NVL(DOME_FL,'%')          like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 9
                sQuery = sQuery + " AND NVL(COUNTRY_CD,'%')       like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  DEST_CD  ASC "
                
            Case "A"  'APPLY_ITEM Code
            
                ssWhere.Col = 1
                sQuery = "          AND APLY_ITEM                 like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(APLY_ITEM_NAME,'%')   like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  APLY_ITEM  ASC "
    
            Case "S"  'Stlgrd Code
            
                ssWhere.Col = 1
                sQuery = " WHERE        STLGRD                    like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(STEEL_GRD_DETAIL,'%') like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  STLGRD  ASC "
                
            Case "T"  'StdSPEC Code
            
                ssWhere.Col = 1
                sQuery = " WHERE        StdSPEC                   like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND StdSPEC_YY                like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  StdSPEC  ASC "
                
            Case "L"  'Melt STD CODE
            
                ssWhere.Col = 1
                sQuery = " WHERE        MLT_STD_NO                like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(APP_DATE,'%')         like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  MILL_STD_NO  ASC "
            
            Case "N"  'Nisco STD CODE
            
                ssWhere.Col = 1
                sQuery = " WHERE        NISCO_QUALITY_NO          like '" & Trim(ssWhere.Text) & "%' "
            
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(APPDATE,'%')          like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  NISCO_QUALITY_NO  ASC "
            
            Case "R"  'Roll STD CODE
            
                ssWhere.Col = 1
                sQuery = " WHERE        MILL_STD_NO               like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = sQuery + " AND NVL(APP_DATE,'%')         like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  MILL_STD_NO  ASC "
            
            Case "V"  'STD DELV CODE
            
                ssWhere.Col = 1
                sQuery = " WHERE        DEV_STD_CD                like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  DEV_STD_CD  ASC "
            
            Case "E"  'Cust STD CODE
            
                ssWhere.Col = 1
                sQuery = " WHERE        CUST_SPEC_NO              like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  CUST_SPEC_NO  ASC "
                
            Case "CHEM"  'CHEMICAL CODE
                
                ssWhere.Col = 1
                sQuery = "WHERE         CHEM_COMP_CD              like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  CHEM_COMP_SEQ  ASC "
                
            Case "G"  'THK GROUP
            
                ssWhere.Col = 1
                sQuery = "          AND THK_CD                    like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = "          AND FR_THK                    like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 3
                sQuery = "          AND TO_THK                    like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY  THK_CD  ASC, FR_THK ASC,  TO_THK ASC "
                
                
            Case "W"  'WID GROUP
            
                ssWhere.Col = 1
                sQuery = "          AND WID_CD                    like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = "          AND FR_WID                    like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 3
                sQuery = "          AND TO_WID                    like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY   WID_CD  ASC, FR_WID ASC,  TO_WID ASC "
                
            Case "RWG"  'ROLL WID GROUP
            
                ssWhere.Col = 1
                sQuery = "          AND WID_GRP_CD                like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = "          AND MINI                      like '" & Trim(ssWhere.Text) & "%' "
                 
                ssWhere.Col = 3
                sQuery = "          AND MAXI                      like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY   WID_GRP_CD  ASC, MINI ASC,  MAXI ASC "
                
            Case "RTG"  'ROLL THK GRUOP
            
                sQuery = "          AND THK_GRP_CD                like '" & Trim(ssWhere.Text) & "%' "
                
                ssWhere.Col = 2
                sQuery = "          AND MINI                      like '" & Trim(ssWhere.Text) & "%' "
                 
                ssWhere.Col = 3
                sQuery = "          AND MAXI                      like '" & Trim(ssWhere.Text) & "%' "
                
                sQuery = sQuery + " ORDER  BY   THK_GRP_CD  ASC, MINI ASC,  MAXI ASC "
                
            Case "HC"  'HEAT CONDITION
            
                ssWhere.Col = 1
                sQuery = " WHERE    HTM_COND                      like '" & Trim(ssWhere.Text) & "%' "
                ssWhere.Col = 2
                sQuery = sQuery + " AND HTM_COND_TXT              like '" & Trim(ssWhere.Text) & "%' "
                  
                sQuery = sQuery + " ORDER  BY  HTM_COND  ASC "
                
    End Select
    
    Call Gf_DD_Display(M_CN1, DD.sQuery + sQuery, False)

End Sub

Private Sub cmd_Selection_Click()

    If ssResult.ActiveRow > 0 Then
    
        ssResult.Row = ssResult.ActiveRow: ssResult.Col = 1
        
        If DD.sWitch = "MS" Then
            DD.rControl.Item(1).Text = ssResult.Text
        Else
            DD.sPname.Col = DD.rControl.Item(1)
            DD.sPname.Text = ssResult.Text
        End If
        
        Select Case DD.DataDicType
        
            Case "M"  'Common Code
            
                Select Case DD.nameType
                    Case "1"            'Short Name
                        ssResult.Col = 2
                    Case "2"            'Full Name
                        ssResult.Col = 3
                    Case "3"            'Short Eng Name
                        ssResult.Col = 4
                    Case "4"            'Full Eng Name
                        ssResult.Col = 5
                End Select
                
            Case "U"  'Order Usage Code
                ssResult.Col = 2
                
            Case "C"  'Customer Code
            
                Select Case DD.nameType
                    Case "1"            'Name
                        ssResult.Col = 2
                    Case "2"            'Eng Name
                        ssResult.Col = 3
                End Select
                
            Case "D"  'Customer Destination Code
                
                Select Case DD.nameType
                    Case "1"            'Name
                        ssResult.Col = 4
                    Case "2"            'Eng Name
                        ssResult.Col = 5
                End Select
                
            Case "A"  'Apply_Item Code
                ssResult.Col = 2
                
            Case "S"  'Stlgrd Code
                ssResult.Col = 2
                
            Case "T"  'StdSPEC Code
                 Select Case DD.rControl.Count
                    Case 2
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                    Case 3
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                        ssResult.Col = 3
                        DD.rControl.Item(6).Text = ssResult.Text
                    Case 4
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                        ssResult.Col = 3
                        DD.rControl.Item(6).Text = ssResult.Text
                        ssResult.Col = 4
                        DD.rControl.Item(5).Text = ssResult.Text
                End Select
                
                ssResult.Col = 2
                
            Case "L"  'Melt STD CODE
                ssResult.Col = 2
            
            Case "N"  'Nisco STD CODE
                ssResult.Col = 2
            
            Case "R"  'Roll STD CODE
                ssResult.Col = 2
            
            Case "V"  'STD DELV CODE
                ssResult.Col = 2
            
            Case "E"  'Cust STD CODE
                ssResult.Col = 2
            
            Case "G"  'THK GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
                
            Case "W"  'WID GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
                
            Case "RTG"  'ROLL THK GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
                
            Case "RWG"  'ROLL WID GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
'2007.11.06 HYS
            Case "HC"  'HEAT TREATMENT CONDITION
                ssResult.Col = 2
                
        End Select
        
        If DD.sWitch = "MS" Then
            
            If DD.rControl.Count > 1 Then
                DD.rControl.Item(2).Text = ssResult.Text
            End If
            
        Else
        
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
            End If
            
            DD.sSelect = True
            
        End If
        
        Unload Me
        
    End If

End Sub

Private Sub Form_Activate()

    ssWhere.Row = 1: ssWhere.Col = 1
    
    If DD.sWitch = "MS" Then
        ssWhere.Text = DD.rControl.Item(1).Text
    Else
        DD.sPname.Col = DD.rControl.Item(1)
        ssWhere.Text = DD.sPname.Text
    End If
    
    Select Case DD.DataDicType
        
        Case "M"  'Common Code
        
            Select Case DD.nameType
                Case "1"            'Short Name
                    ssWhere.Col = 2
                Case "2"            'Full Name
                    ssWhere.Col = 3
                Case "3"            'Short Eng Name
                    ssWhere.Col = 4
                Case "4"            'Full Eng Name
                    ssWhere.Col = 5
            End Select
            
        Case "U"  'Order Usage Code
            ssWhere.Col = 2
            
        Case "C"  'Customer Code
        
            Select Case DD.nameType
                Case "1"            'Name
                    ssWhere.Col = 2
                Case "2"            'Eng Name
                    ssWhere.Col = 3
            End Select
        
        Case "D"  'Customer Destination Code
            
            Select Case DD.nameType
                Case "1"            'Name
                    ssWhere.Col = 4
                Case "2"            'Eng Name
                    ssWhere.Col = 5
            End Select
            
        Case "A"  'Apply_Item Code
            ssWhere.Col = 2
            
        Case "S"  'Stlgrd Code
            ssWhere.Col = 2
            
        Case "T"  'StdSPEC Code
            ssWhere.Col = 2
            
        Case "L"  'Melt STD CODE
            ssWhere.Col = 2
            
        Case "N"  'Nisco STD CODE
            ssWhere.Col = 2
        
        Case "R"  'Roll STD CODE
            ssWhere.Col = 2
        
        Case "V"  'STD DELV CODE
            ssWhere.Col = 2
        
        Case "E"  'Cust STD CODE
            ssWhere.Col = 2
            
        Case "G"  'THK GROUP CODE
            ssWhere.Col = 2
            
        Case "W"  'WID GROUP CODE
            ssWhere.Col = 2
            
        Case "RTG"  'ROLL THK GROUP CODE
            ssWhere.Col = 2
            
        Case "RWG"  'ROLL WID GROUP CODE
            ssWhere.Col = 2
            
'2007.11.06 HYS
         Case "HC"  'HEAT TREATMENT CONDITION
            ssWhere.Col = 2
    End Select
    
    If DD.sWitch = "MS" Then
    
        If DD.rControl.Count > 1 Then
            ssWhere.Text = DD.rControl.Item(2).Text
        End If
        
    Else
    
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            ssWhere.Text = DD.sPname.Text
        End If
        
    End If

    Call ssWhere_setting
    Call ssResult_setting

    Call Gp_Sp_ColGet(ssWhere, "Z-System.INI", Me.Name, DD.DataDicType)
    Call Gp_Sp_ColGet(ssResult, "Z-System.INI", Me.Name, DD.DataDicType)
    
    Me.BackColor = &HE0E0E0
    
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

    Call Gp_FormLoc_Get(Me, DD.DataDicType)
    PrevRow = 0
    Me.KeyPreview = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_FormLoc_Set(Me, DD.DataDicType)
    
    Call Gp_Sp_ColSet(ssWhere, "Z-System.INI", Me.Name, DD.DataDicType)
    Call Gp_Sp_ColSet(ssResult, "Z-System.INI", Me.Name, DD.DataDicType)
    
    DD.DataDicType = ""
    DD.DicRefType = ""
    DD.nameType = ""
    DD.sQuery = ""
    'DD.sWitch = ""
    DD.sWhere = ""
    DD.sKey = ""
    
    Set DD.wControl = Nothing
    'Set DD.rControl = Nothing
    'Set DD.sPname = Nothing
    
End Sub

'================================='
' Where Spread Setting
'================================='
Public Sub ssWhere_setting()

    With ssWhere
    
        .RowHeight(-1) = 12.54
        .RowHeight(0) = 16
        
        .ColWidth(0) = 4.3
        
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
        
        .ColWidth(0) = 4.3
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .BlockMode = False
        
        .UserResize = UserResizeColumns
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HA4FDE2
        
        .OperationMode = OperationModeRead
        
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
    If Me.Height - 2250 > 1 Then
        pnl_result.Height = Me.Height - 2250
    End If
    
    'Spread Width Setting (ssWhere)
    If Me.ScaleWidth - ssWhere.Left > 1 Then
        ssWhere.Width = Me.ScaleWidth - 180
    End If

    'Spread Width Setting (ssResult)
    If Me.ScaleWidth - ssResult.Left > 1 Then
        ssResult.Width = Me.ScaleWidth - 180
    End If
    
    'Spread Height Setting (ssResult)
    'If Me.ScaleHeight - ssResult.Top > 1 Then
    If pnl_result.Height - 100 > 1 Then
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
    
        ssResult.Row = Row: ssResult.Col = 1
        
        If DD.sWitch = "MS" Then
            DD.rControl.Item(1).Text = ssResult.Text
        Else
            DD.sPname.Col = DD.rControl.Item(1)
            DD.sPname.Text = ssResult.Text
        End If

        Select Case DD.DataDicType
        
            Case "M"  'Common Code
            
                Select Case DD.nameType
                    Case "1"            'Short Name
                        ssResult.Col = 2
                    Case "2"            'Full Name
                        ssResult.Col = 3
                    Case "3"            'Short Eng Name
                        ssResult.Col = 4
                    Case "4"            'Full Eng Name
                        ssResult.Col = 5
                End Select
                
            Case "U"  'Order Usage Code
                ssResult.Col = 2
                
            Case "C"  'Customer Code
            
                Select Case DD.nameType
                    Case "1"            'Name
                        ssResult.Col = 2
                    Case "2"            'Eng Name
                        ssResult.Col = 3
                End Select
                                        
            Case "D"  'Customer Destination Code
                
                Select Case DD.nameType
                    Case "1"            'Name
                        ssResult.Col = 4
                    Case "2"            'Eng Name
                        ssResult.Col = 5
                End Select
        
            Case "A"  'Apply_Item Code
                ssResult.Col = 2
        
            Case "S"  'Stlgrd Code
                ssResult.Col = 2
        
            Case "T"  'StdSPEC Code
                
                Select Case DD.rControl.Count
                    Case 2
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                    Case 3
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                        ssResult.Col = 6
                        DD.rControl.Item(3).Text = ssResult.Text
                    Case 4
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                        ssResult.Col = 6
                        DD.rControl.Item(3).Text = ssResult.Text
                        ssResult.Col = 5
                        DD.rControl.Item(4).Text = ssResult.Text
                End Select
                
                ssResult.Col = 2
                
            Case "L"  'Melt STD CODE
                ssResult.Col = 2
            
            Case "N"  'Nisco STD CODE
                Select Case DD.rControl.Count
                    Case 2
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                    Case 3
                        ssResult.Col = 2
                        DD.rControl.Item(2).Text = ssResult.Text
                        ssResult.Col = 3
                        DD.rControl.Item(3).Text = ssResult.Text
                End Select
                
                ssResult.Col = 2
            
            Case "R"  'Roll STD CODE
                ssResult.Col = 2
            
            Case "V"  'STD DELV CODE
                ssResult.Col = 2
            
            Case "E"  'Cust STD CODE
                ssResult.Col = 2
                
            Case "G"  'THK GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
                
            Case "W"  'WID GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
                
            Case "RTG"  'ROLL THK GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
                
            Case "RWG"  'ROLL WID GROUP Code
                ssResult.Col = 2
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 3
                DD.sPname.Col = DD.rControl.Item(3)
                DD.sPname.Text = ssResult.Text
                ssResult.Col = 2
                
' 2007.11.6 HYS
            Case "HC"  'HEAT TREATMENT CONDITION
                ssResult.Col = 2
' 2008.1.2 Kim.Sung.Ho
            Case "EMP"  'EMPLOYEE
                ssResult.Col = 2
                
        End Select
        
        If DD.sWitch = "MS" Then
        
            If DD.rControl.Count > 1 Then
                DD.rControl.Item(2).Text = ssResult.Text
            End If
            
        Else
        
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
            End If
            
            DD.sSelect = True
            
        End If
        
        Unload Me
        
    End If
    
End Sub

Private Sub ssResult_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    
    If KeyCode = vbKeyReturn Then

        ssResult.Row = ssResult.ActiveRow: ssResult.Col = 1
        
        If DD.sWitch = "MS" Then
            DD.rControl.Item(1).Text = ssResult.Text
        Else
            DD.sPname.Col = DD.rControl.Item(1)
            DD.sPname.Text = ssResult.Text
        End If
        
        Select Case DD.DataDicType
        
            Case "M"  'Common Code
            
                Select Case DD.nameType
                    Case "1"            'Short Name
                        ssResult.Col = 2
                    Case "2"            'Full Name
                        ssResult.Col = 3
                    Case "3"            'Short Eng Name
                        ssResult.Col = 4
                    Case "4"            'Full Eng Name
                        ssResult.Col = 5
                End Select
                
            Case "U"  'Order Usage Code
                ssResult.Col = 2
                
            Case "C"  'Customer Code
            
                Select Case DD.nameType
                    Case "1"            'Name
                        ssResult.Col = 2
                    Case "2"            'Eng Name
                        ssResult.Col = 3
                End Select
            
            Case "D"  'Customer Destination Code
                
                Select Case DD.nameType
                    Case "1"            'Name
                        ssResult.Col = 4
                    Case "2"            'Eng Name
                        ssResult.Col = 5
                End Select
                
            Case "A"  'Apply_Item Code
                ssResult.Col = 2
            
            Case "S"  'Stlgrd Code
                ssResult.Col = 2
            
            Case "T"  'StdSPEC Code
                ssResult.Col = 2
            
            Case "L"  'Melt STD CODE
                ssWhere.Col = 1
                ssWhere.Col = 2
                
            Case "N"  'Nisco STD CODE
                ssWhere.Col = 2
                
            Case "R"  'Roll STD CODE
                ssWhere.Col = 2
                
            Case "V"  'STD DELV CODE
                ssWhere.Col = 2
                
            Case "E"  'Cust STD CODE
                ssWhere.Col = 2
                
            Case "G"  'THK GROUP CODE
                ssWhere.Col = 2
                
            Case "W"  'WID GROUP CODE
                ssWhere.Col = 2
                
            Case "RTG"  'ROLL THK GROUP CODE
                ssWhere.Col = 2
                
            Case "RWG"  'ROLL WID GROUP CODE
                ssWhere.Col = 2
                
'2007.11.06 HYS
            Case "HC"  'HEAT TRAETMENT CONDITION
                ssWhere.Col = 2
            
        End Select
        
        If DD.sWitch = "MS" Then
            
            If DD.rControl.Count > 1 Then
                DD.rControl.Item(2).Text = ssResult.Text
            End If
            
        Else
        
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                DD.sPname.Text = ssResult.Text
            End If
            
            DD.sSelect = True
            
        End If
        
        Unload Me
        
    End If

End Sub

Private Sub ssResult_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Col As Long, Row As Long

    ssResult.GetCellFromScreenCoord Col, Row, x, y

    If Row <= 0 Or PrevRow = Row Then Exit Sub
    
    Call Gp_Sp_RowColor(ssResult, Row, , &HA4FDE2)
    Call Gp_Sp_RowColor(ssResult, PrevRow)
    
    PrevRow = Row
    
End Sub

Private Sub ssWhere_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)
    
    Dim iCol As Integer
    
    For iCol = 1 To ssWhere.MaxCols
        ssResult.ColWidth(iCol) = ssWhere.ColWidth(iCol)
    Next iCol
    
End Sub
