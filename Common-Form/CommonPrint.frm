VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form CommonPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Preview"
   ClientHeight    =   9210
   ClientLeft      =   180
   ClientTop       =   480
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CommonPrint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9210
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboZoom 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   90
      Width           =   2265
   End
   Begin FPSpread.vaSpreadPreview prnPreview 
      Height          =   8595
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   15255
      _Version        =   393216
      _ExtentX        =   26908
      _ExtentY        =   15161
      _StockProps     =   96
      BorderStyle     =   1
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
   Begin Threed.SSCommand cmdMenu 
      Height          =   495
      Index           =   0
      Left            =   13770
      TabIndex        =   1
      Top             =   60
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   873
      _Version        =   196609
      Font3D          =   3
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CommonPrint.frx":030A
      Caption         =   "&Exit    "
      Alignment       =   4
      PictureAlignment=   1
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmdMenu 
      Height          =   495
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   873
      _Version        =   196609
      Font3D          =   3
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CommonPrint.frx":0DDC
      Caption         =   "&Print   "
      Alignment       =   4
      PictureAlignment=   1
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmdMenu 
      Height          =   510
      Index           =   2
      Left            =   1530
      TabIndex        =   3
      Top             =   60
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   900
      _Version        =   196609
      Font3D          =   3
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CommonPrint.frx":1AD6
      Caption         =   "&Setup   "
      Alignment       =   4
      PictureAlignment=   1
      BevelWidth      =   1
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdMenu 
      Height          =   495
      Index           =   3
      Left            =   7470
      TabIndex        =   4
      Top             =   60
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   873
      _Version        =   196609
      Font3D          =   3
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CommonPrint.frx":1F28
      Caption         =   "      &Next"
      Alignment       =   1
      PictureAlignment=   4
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmdMenu 
      Height          =   495
      Index           =   4
      Left            =   3330
      TabIndex        =   5
      Top             =   60
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   873
      _Version        =   196609
      Font3D          =   3
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CommonPrint.frx":237A
      Caption         =   "&Previous "
      Alignment       =   4
      PictureAlignment=   1
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmdMenu 
      Height          =   495
      Index           =   5
      Left            =   9480
      TabIndex        =   6
      Top             =   60
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   873
      _Version        =   196609
      Font3D          =   3
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CommonPrint.frx":27CC
      Caption         =   "&Zoom   "
      Alignment       =   4
      PictureAlignment=   1
      BevelWidth      =   1
   End
   Begin InDate.ULabel ULabel1 
      Height          =   465
      Left            =   4800
      Top             =   60
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   820
      Caption         =   "ADFAS"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "CommonPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboZoom_Click()
    zoomIndex = cboZoom.ListIndex
    Call GetZoom(zoomIndex)
End Sub

Private Sub cmdMenu_Click(Index As Integer)
    On Error GoTo Err
    
    With prnPreview
        Select Case Index
            Case 0  'exit
                Set prnSp = Nothing
                Unload Me
            Case 1  'print
                prnSp.PrintSheet
                Call cmdMenu_Click(0)
            Case 2  'setup
                PageSetup.Show vbModal
            Case 3  'next
                If .PageCurrent < prnSp.PrintPageCount Then
                    .PageCurrent = .PageCurrent + .PagesPerScreen
                    
                    cmdMenu(3).Enabled = True
                    cmdMenu(4).Enabled = True
                End If
                
                 'If at last page, disable button
                If .PageCurrent >= prnSp.PrintPageCount - .PagesPerScreen Then
                    cmdMenu(3).Enabled = False
                End If
            Case 4  'previous
                If .PageCurrent > 1 Then
                    .PageCurrent = .PageCurrent - .PagesPerScreen
                    cmdMenu(3).Enabled = True
                    cmdMenu(4).Enabled = True
                End If
                
                'If at first page, disable button
                If .PageCurrent = 1 Then
                    cmdMenu(4).Enabled = False
                End If
            Case 5
                .ZoomState = 3
        End Select
    End With
    
    Exit Sub
Err:
    MsgBox Err.Description & "(" & Err.Number & ")", vbCritical, "Print"
    
End Sub

Private Sub Form_Activate()
    'Attach preview control to Spread
    prnPreview.hWndSpread = prnSp.hwnd
   
    'Update page count listing
    UpdatePageCount
    
    With prnPreview
        If .PageCurrent = 1 Then
            cmdMenu(4).Enabled = False
        Else
            cmdMenu(4).Enabled = True
        End If
        
        If .PageCurrent = prnSp.PrintPageCount Then
            cmdMenu(3).Enabled = False
        Else
            cmdMenu(3).Enabled = True
        End If
    End With
End Sub

Private Sub Form_Load()
'    GF_LoadPicture Image1
    
    'Disable Previous button
    cmdMenu(4).Enabled = False
        
    'Get the zoom display
    GetZoom zoomIndex
    
    'Set up page numbering
    If prnSp.PrintPageCount = 1 Then
        'Disable Next button if only one page
        cmdMenu(3).Enabled = False
    End If
    
    'Populate Zooming combobox
    With cboZoom
        .AddItem "200%"
        .AddItem "150%"
        .AddItem "100%"
        .AddItem "75%"
        .AddItem "50%"
        .AddItem "25%"
        .AddItem "10%"
        .AddItem "Page Width"
        .AddItem "Page Height"
        .AddItem "Whole Page"
        .AddItem "Two Pages"
        .AddItem "Three Pages"
        .AddItem "Four Pages"
        .AddItem "Six Pages"
        
        'Get the zoom display
        .ListIndex = zoomIndex
    End With
End Sub

Private Sub Form_Resize()
    prnPreview.Move 0, Image1.Height, ScaleWidth, ScaleHeight - Image1.Height
End Sub

Sub UpdatePageCount()
    'Page Count
    ULabel1.Caption = "Page " & prnPreview.PageCurrent & " of " & prnSp.PrintPageCount
    
End Sub

Private Sub prnPreview_PageChange(ByVal Page As Long)
    UpdatePageCount
End Sub

