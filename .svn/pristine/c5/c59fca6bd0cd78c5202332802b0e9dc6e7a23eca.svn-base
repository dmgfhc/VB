VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PageSetup 
   Caption         =   "Setup"
   ClientHeight    =   3765
   ClientLeft      =   4275
   ClientTop       =   2235
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3765
   ScaleWidth      =   8400
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   0
      TabIndex        =   28
      Top             =   3090
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1138
      _Version        =   196609
      BevelOuter      =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdCancel 
         Height          =   405
         Left            =   6720
         TabIndex        =   29
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   714
         _Version        =   196609
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   405
         Left            =   5040
         TabIndex        =   30
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   714
         _Version        =   196609
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OK"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand cmdPrintSetUp 
         Height          =   405
         Left            =   150
         TabIndex        =   31
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   714
         _Version        =   196609
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Printer Setup"
         BevelWidth      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4170
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin Threed.SSFrame fraOptions 
      Height          =   1665
      Left            =   4230
      TabIndex        =   20
      Top             =   1350
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2937
      _Version        =   196609
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ Print Options ]"
      Begin VB.CheckBox Check1 
         Caption         =   "Column Headers"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   390
         Value           =   1  '»Æ¿Œ
         Width           =   1965
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Row Headers"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   690
         Value           =   1  '»Æ¿Œ
         Width           =   1965
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Grid Lines"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   990
         Value           =   1  '»Æ¿Œ
         Width           =   1425
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Border"
         Height          =   255
         Index           =   3
         Left            =   2250
         TabIndex        =   24
         Top             =   390
         Value           =   1  '»Æ¿Œ
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Data Cells Only"
         Height          =   255
         Index           =   4
         Left            =   2250
         TabIndex        =   23
         Top             =   690
         Value           =   1  '»Æ¿Œ
         Width           =   1785
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Color"
         Height          =   255
         Index           =   5
         Left            =   2250
         TabIndex        =   22
         Top             =   990
         Value           =   1  '»Æ¿Œ
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Shadows"
         Height          =   255
         Index           =   6
         Left            =   2250
         TabIndex        =   21
         Top             =   1290
         Value           =   1  '»Æ¿Œ
         Width           =   1365
      End
   End
   Begin Threed.SSFrame fraRange 
      Height          =   1665
      Left            =   60
      TabIndex        =   12
      Top             =   1350
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   2937
      _Version        =   196609
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ Page Range ]"
      Begin VB.OptionButton Option1 
         Caption         =   "All"
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   390
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Selected Cells"
         Height          =   225
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   660
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Current Page"
         Height          =   225
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   930
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pages"
         Height          =   225
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   1230
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'ø¿∏•¬  ∏¬√„
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1740
         TabIndex        =   14
         Text            =   "1"
         Top             =   1200
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'ø¿∏•¬  ∏¬√„
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   13
         Text            =   "1"
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   1230
         Width           =   195
      End
   End
   Begin Threed.SSFrame fraMargins 
      Height          =   1155
      Left            =   4230
      TabIndex        =   3
      Top             =   90
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2037
      _Version        =   196609
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ Page Margins (inch) ]"
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   4
         Top             =   390
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Format          =   "0.00"
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   5
         Top             =   690
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Format          =   "0.00"
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   390
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Format          =   "0.00"
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   7
         Top             =   690
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Format          =   "0.00"
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'ø¿∏•¬  ∏¬√„
         Caption         =   "Left:"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   11
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'ø¿∏•¬  ∏¬√„
         Caption         =   "Right:"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   750
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'ø¿∏•¬  ∏¬√„
         Caption         =   "Top:"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   9
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'ø¿∏•¬  ∏¬√„
         Caption         =   "Bottom:"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2070
         TabIndex        =   8
         Top             =   690
         Width           =   645
      End
   End
   Begin Threed.SSFrame fraOrientation 
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   2037
      _Version        =   196609
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ Page Orientation ]"
      Begin VB.OptionButton porientation 
         Caption         =   "Portrait"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   540
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton porientation 
         Caption         =   "Landscape"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2610
         TabIndex        =   1
         Top             =   540
         Width           =   1365
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "PageSetup.frx":0000
         Top             =   420
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   1
         Left            =   2010
         Picture         =   "PageSetup.frx":0AC2
         Top             =   420
         Width           =   495
      End
   End
End
Attribute VB_Name = "PageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'OK button
    With prnSp
        'Update margins
        .PrintMarginTop = CDbl(Replace(pagemargin(2).Text, "_", "")) * 1440
        .PrintMarginBottom = CDbl(Replace(pagemargin(3).Text, "_", "")) * 1440
        .PrintMarginLeft = CDbl(Replace(pagemargin(0).Text, "_", "")) * 1440
        .PrintMarginRight = CDbl(Replace(pagemargin(1).Text, "_", "")) * 1440
        
        'Change the page orientation
        If porientation(0).Value = True Then
            'Portrait
            .PrintOrientation = PrintOrientationPortrait
        Else
            'Landscape
            .PrintOrientation = PrintOrientationLandscape
        End If
        
         'Set printing options for spreadsheet
         .PrintColHeaders = Check1(0).Value
         .PrintRowHeaders = Check1(1).Value
         .PrintBorder = Check1(3).Value
         .PrintColor = Check1(5).Value
         .PrintGrid = Check1(2).Value
         .PrintShadows = Check1(6).Value
         .PrintUseDataMax = Check1(4).Value
        
         'Page Range
         If Option1(0).Value = True Then
             'All
             .PrintType = SS_PRINT_ALL
         ElseIf Option1(1).Value = True Then
            'Selected cells
             .Col = .SelBlockCol
             .Col2 = .SelBlockCol2
             .Row = .SelBlockRow
             .Row2 = .SelBlockRow2
             .PrintType = SS_PRINT_CELL_RANGE
         ElseIf Option1(2).Value = True Then
             'Current Page
             .PrintType = SS_PRINT_CURRENT_PAGE
         Else
             'Pages
             .PrintPageStart = CInt(Text1(0).Text)
             .PrintPageEnd = CInt(Text1(1).Text)
             .PrintType = SS_PRINT_PAGE_RANGE
         End If
    End With
    
    Call cmdCancel_Click
End Sub

Private Sub cmdPrintSetUp_Click()
    CommonDialog1.ShowPrinter
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    With prnSp
        'Get page margins (convert to inches) and format
        pagemargin(0).Text = Format(.PrintMarginLeft / 1440, "0.00")
        pagemargin(1).Text = Format(.PrintMarginRight / 1440, "0.00")
        pagemargin(2).Text = Format(.PrintMarginTop / 1440, "0.00")
        pagemargin(3).Text = Format(.PrintMarginBottom / 1440, "0.00")
        
        'Get page orientation
        If .PrintOrientation = PrintOrientationLandscape Then
            porientation(1) = True
        Else
            porientation(0) = True
        End If

         'Set printing options for spreadsheet
         For i = 0 To 6
            Check1(i).Value = vbUnchecked
         Next i
         
         If .PrintColHeaders Then Check1(0).Value = vbChecked
         If .PrintRowHeaders Then Check1(1).Value = vbChecked
         If .PrintBorder Then Check1(3).Value = vbChecked
         If .PrintColor Then Check1(5).Value = vbChecked
         If .PrintGrid Then Check1(2).Value = vbChecked
         If .PrintShadows Then Check1(6).Value = vbChecked
         If .PrintUseDataMax Then Check1(4).Value = vbChecked
    End With
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 3 Then
        Text1(0).Enabled = True
        Text1(1).Enabled = True
        Text1(0).SetFocus
    Else
        Text1(0).Enabled = False
        Text1(1).Enabled = False
    End If
End Sub

