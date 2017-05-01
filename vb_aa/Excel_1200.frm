VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form frm_Excel 
   Caption         =   "选择Excel文件导入_frm_Excel"
   ClientHeight    =   4395
   ClientLeft      =   8340
   ClientTop       =   4650
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5055
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txt_load_file 
      Height          =   270
      Left            =   3090
      TabIndex        =   8
      Top             =   255
      Visible         =   0   'False
      Width           =   1980
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   2430
      Left            =   120
      TabIndex        =   5
      Top             =   510
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4286
      _Version        =   196609
      PaneTree        =   "Excel_1200.frx":0000
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   2730
         Pattern         =   "*.xls"
         TabIndex        =   7
         Top             =   30
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   2610
      End
   End
   Begin VB.Frame framProcess 
      BackColor       =   &H00E0E0E0&
      Caption         =   "文件加载中......"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   810
      Left            =   120
      TabIndex        =   3
      Top             =   3015
      Visible         =   0   'False
      Width           =   4815
      Begin MSComctlLib.ProgressBar ProgBar 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin Threed.SSCommand cmd_ok 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3915
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
   End
   Begin Threed.SSCommand cmd_cancel 
      Height          =   375
      Left            =   3090
      TabIndex        =   2
      Top             =   3915
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
   End
End
Attribute VB_Name = "frm_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()

    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim iCount          As Integer
    Dim iRow            As Integer
    Dim iCnt            As Integer
    Dim iXrow           As Integer
    Dim iXCnt           As Integer
    Dim iDR             As Integer
    Dim sPath           As String
    
    On Error GoTo ErrProc
    
    If Trim(File1.FileName) = "" Then
        MsgBox "请选择要导入的Excel文件! ", vbCritical + vbOKOnly, "系统提示信息"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    framProcess.Visible = True
    
    Set xlApp = GetObject("", "Excel.Application")
    
    If Err.Number = 429 Then
        Set xlApp = CreateObject("", "Excel.Application")
    End If
    
    sPath = Dir1.Path
    
    If Mid(sPath, Len(sPath), 1) = "\" Then
        sPath = Mid(sPath, 1, Len(sPath) - 1)
    End If
    
    xlApp.Workbooks.Open (Trim(sPath & "\" & File1.FileName))
    
    Set xlSheet = xlApp.Worksheets(1)
    
    iDR = 0
    iXCnt = 0
    iXrow = 2
    
    While CStr(xlSheet.Cells(iXrow, 1)) > " "
        iXCnt = iXCnt + 1
        iXrow = iXrow + 1
    Wend
       
    ProgBar.Min = 0
    ProgBar.Max = iXCnt
            
    For iRow = 1 To iXCnt
        
        ProgBar.Value = iRow
        framProcess.Caption = "正在读取Excel文件!" & CStr(iRow) & " / " & CStr(iXCnt)
        DoEvents
        Select Case txt_load_file.Text
               ''销售计划导入
               Case "AAA1200C"
                With AAA1200C.ss2
                    .MaxRows = iRow
                    .Row = iRow
                    .Col = 0:       .Text = "Input"
                    For iCol = 1 To .MaxCols - 4
                        .Col = iCol
                        .Text = CStr(xlSheet.Cells(iRow + 1, iCol))
                        If .Col = .MaxCols - 4 Then
                           .BlockMode = False
                           .Lock = False
                           .BackColor = &HC0FFFF
                        End If
                    Next iCol
                    .Col = .MaxCols - 2:       .Text = sUserID   'INS_EMP
                End With
               ''技术参数导入
               Case "AAA1020C"
                With AAA1020C.ss2
                    .MaxRows = iRow
                    .Row = iRow
                    .Col = 0:       .Text = "Input"
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                        If .Col = 2 Then
                            .Text = Format(CStr(xlSheet.Cells(iRow + 1, iCol)), "000")
                        Else
                            .Text = CStr(xlSheet.Cells(iRow + 1, iCol))
                        End If
                        
                        If .Col = .MaxCols - 2 Then
                           .BlockMode = False
                           .Lock = False
                           .BackColor = &HC0FFFF
                        End If
                    Next iCol
                    .Col = .MaxCols:       .Text = sUserID   'INS_EMP
                End With
               ''板坯计划导入
               Case "AAA1070C"
                With AAA1070C.ss2
                    .MaxRows = iRow
                    .Row = iRow
                    .Col = 0:       .Text = "Input"
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                        If .Col = 3 Then
                            .Text = Format(CStr(xlSheet.Cells(iRow + 1, iCol)), "000")
                        Else
                            .Text = CStr(xlSheet.Cells(iRow + 1, iCol))
                        End If
                        
                        If .Col = .MaxCols - 4 Then
                           .BlockMode = False
                           .Lock = False
                           .BackColor = &HC0FFFF
                        End If
                    Next iCol
                    .Col = .MaxCols - 2:   .Text = sUserID   'INS_EMP
                End With
        End Select
    Next iRow
            
    xlApp.ActiveWorkbook.Close Trim(File1.FileName)
    
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlApp = Nothing
    
    Screen.MousePointer = vbDefault
    
    Unload Me
    
    Exit Sub

ErrProc:

    If Err.Number = 429 Then
        MsgBox "Microsoft Excel Program Not Installed"
    Else
        MsgBox Err.Number & Err.Description
    End If
    
    Set xlSheet = Nothing
    xlApp.ActiveWorkbook.Close False
    xlApp.Quit
    Set xlApp = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub File1_DblClick()
Call Cmd_Ok_Click
End Sub

Private Sub Form_Load()
Drive1.Drive = "C:/"
End Sub
