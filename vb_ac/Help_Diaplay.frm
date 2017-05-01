VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form HelpDiaplay 
   Caption         =   "界面说明书"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15240
   Begin Threed.SSFrame sFrame 
      Height          =   9195
      Left            =   2910
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   16219
      _Version        =   196609
      BackColor       =   14737632
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   14737632
         Caption         =   "1"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   2
         Left            =   525
         TabIndex        =   5
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "2"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   3
         Left            =   870
         TabIndex        =   6
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "3"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   4
         Left            =   1215
         TabIndex        =   7
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "4"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   5
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "5"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   6
         Left            =   1905
         TabIndex        =   9
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "6"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   7
         Left            =   2250
         TabIndex        =   10
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "7"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   8
         Left            =   2595
         TabIndex        =   11
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "8"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   9
         Left            =   2940
         TabIndex        =   12
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "9"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_img_index 
         Height          =   330
         Index           =   10
         Left            =   3285
         TabIndex        =   13
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   14737632
         Caption         =   "10"
         ButtonStyle     =   4
      End
      Begin VB.Image Image1 
         Height          =   8520
         Left            =   120
         Top             =   495
         Width           =   12000
      End
   End
   Begin Threed.SSCommand cmd_Image 
      Height          =   300
      Left            =   13560
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   196609
      ForeColor       =   16711680
      Caption         =   "Open Image"
   End
   Begin VB.TextBox txt_Detail 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9630
      Left            =   2910
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "Help_Diaplay.frx":0000
      Top             =   165
      Width           =   12240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15
      Top             =   9360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help_Diaplay.frx":0006
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help_Diaplay.frx":0458
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help_Diaplay.frx":08AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help_Diaplay.frx":0CFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tre_view 
      Height          =   9630
      Left            =   165
      TabIndex        =   1
      Top             =   165
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   16986
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "HelpDiaplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPgmId          As String
Dim iImageCnt       As Integer
Dim sServerIP       As String
Dim sServerID       As String
Dim sServerPWD      As String
Dim sServerPATH     As String


Private Sub cmd_exit_Click()
    sFrame.Visible = False
End Sub

Private Sub cmd_Image_Click()

    Dim sQuery      As String
    Dim AdoRs       As ADODB.Recordset
    
    On Error Resume Next
    
    If sFrame.Visible = True Then
        sFrame.Visible = False
        cmd_Image.Caption = "Open Image"
        Exit Sub
    Else
        sFrame.Visible = True
        cmd_Image.Caption = "Close Image"
    End If
        
    If sServerIP = "" Then
        Set AdoRs = New ADODB.Recordset
    
        sQuery = "SELECT SERVER_IP, SERVER_ID, SERVER_PWD, SERVER_PATH FROM ZP_SERVERINFO "
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        If Not AdoRs.BOF And Not AdoRs.EOF Then
    
            If VarType(AdoRs.Fields(0)) = vbNull Then
                sServerIP = ""
            Else
                sServerIP = AdoRs.Fields(0)
            End If
    
            If VarType(AdoRs.Fields(1)) = vbNull Then
                sServerID = ""
            Else
                sServerID = AdoRs.Fields(1)
            End If
    
            If VarType(AdoRs.Fields(2)) = vbNull Then
                sServerPWD = ""
            Else
                sServerPWD = AdoRs.Fields(2)
            End If
    
            If VarType(AdoRs.Fields(3)) = vbNull Then
                sServerPATH = ""
            Else
                sServerPATH = AdoRs.Fields(3)
            End If
    
        End If
    
        AdoRs.Close
        Set AdoRs = Nothing
    End If
    
    Call cmd_img_index_Click(1)
End Sub

Private Sub cmd_img_index_Click(Index As Integer)
    Dim iDx As Integer
    
    For iDx = 1 To 10
        cmd_img_index(iDx).ForeColor = &H80000012
    Next iDx
    
    Call Image_Display(Index)
    cmd_img_index(Index).ForeColor = &HFF0000
End Sub

Private Sub Form_Activate()
    Dim sPgmId      As String
    Dim sPgmName    As String
    Dim sQuery      As String
    
    sPgmId = UCase(Me.Tag)
    
    If sPgmId <> "" Then
        
        sQuery = " SELECT    PGMNAME   FROM   ZP_PGMID " & vbCrLf
        sQuery = sQuery & "  WHERE PGMID   = '" & sPgmId & "' " & vbCrLf
        
        Set AdoRs = Gf_Ms_Rset(M_CN1, sQuery)
        If Not AdoRs.EOF Then
            sPgmName = Trim(AdoRs.Fields(0))
        End If
    End If
    
    Call TreeView_Setting(sPgmName)
    txt_Detail.Locked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HelpDiaplay.Tag = ""
    Unload Me
End Sub


Private Sub TreeView_Setting(sPgmName As String)
    Dim sQuery  As String
    Dim AdoRs   As New ADODB.Recordset
    Dim nodX    As Node
    Dim iLoc    As Integer
    Dim iDx     As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    Set AdoRs = New ADODB.Recordset
    
    iLoc = InStr(1, MDIMain.Caption, " ")
    If iLoc = 0 Then iLoc = Len(MDIMain.Caption)
    
    Set nodX = tre_view.Nodes.Add(, , "root", Left(MDIMain.Caption, iLoc), 1, 2)
    nodX.Image = 1
    
    sQuery = "   SELECT      PGMID,                     " & vbCrLf
    sQuery = sQuery & "      PGMNAME                    " & vbCrLf
    sQuery = sQuery & " FROM ZP_PGMID                   " & vbCrLf
    sQuery = sQuery & "WHERE BIZ_AREA   = '" & MDIMain.Tag & "' " & vbCrLf
    sQuery = sQuery & "  AND HELP_YN    = 'Y'           " & vbCrLf
    sQuery = sQuery & "ORDER BY PGMID "
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.EOF Then
        Do Until AdoRs.EOF Or AdoRs.BOF
            Set nodX = tre_view.Nodes.Add("root", tvwChild, Trim(AdoRs.Fields(0)), AdoRs.Fields(1), 3, 4)
            nodX.Image = 3
            AdoRs.MoveNext
        Loop
        
        txt_Detail = "界面说明书 : "
        
        For iDx = 1 To AdoRs.RecordCount + 1
            If sPgmName = tre_view.Nodes.Item(iDx) Then
                tre_view.Nodes.Item(iDx).Selected = True
                Call tre_view_Click
                Exit For
            End If
        Next iDx
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub sFrame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sFrame.Move sFrame.Left + X, sFrame.Top + Y
End Sub

Private Sub tre_view_Click()
    
    sPgmId = Trim(tre_view.SelectedItem.Key)
    
    Call HelpDisplay
End Sub

Private Sub HelpDisplay()
    Dim iDx         As Integer
    Dim sQuery      As String
    Dim AdoRs       As ADODB.Recordset
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    Set AdoRs = New ADODB.Recordset

    sQuery = "   SELECT      PGMNAME,   PGMID,  " & vbCrLf
    sQuery = sQuery & "      TITLE1,    HELP1,  " & vbCrLf
    sQuery = sQuery & "      TITLE2,    HELP2,  " & vbCrLf
    sQuery = sQuery & "      TITLE3,    HELP3,  " & vbCrLf
    sQuery = sQuery & "      TITLE4,    HELP4,  " & vbCrLf
    sQuery = sQuery & "      IMAGE_CNT          " & vbCrLf
    sQuery = sQuery & " FROM ZP_PGMID                   " & vbCrLf
    sQuery = sQuery & "WHERE PGMID      = '" & sPgmId & "'" & vbCrLf
    sQuery = sQuery & "  AND HELP_YN    = 'Y'           " & vbCrLf
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not AdoRs.BOF And Not AdoRs.EOF Then

        txt_Detail = "界面说明书 : " & AdoRs.Fields(0) & "(" & AdoRs.Fields(1) & ")" & vbCrLf & vbCrLf
        
        If Trim(AdoRs.Fields(2) & "") <> "" Then
            txt_Detail = txt_Detail & "** " & Trim(AdoRs.Fields(2) & "") & vbCrLf
        End If
        txt_Detail = txt_Detail & AdoRs.Fields(3) & vbCrLf & vbCrLf
        
        If Trim(AdoRs.Fields(4) & "") <> "" Then
            txt_Detail = txt_Detail & "** " & Trim(AdoRs.Fields(4) & "") & vbCrLf
        End If
        txt_Detail = txt_Detail & AdoRs.Fields(5) & vbCrLf & vbCrLf
        
        If Trim(AdoRs.Fields(6) & "") <> "" Then
            txt_Detail = txt_Detail & "** " & Trim(AdoRs.Fields(6) & "") & vbCrLf
        End If
        txt_Detail = txt_Detail & AdoRs.Fields(7) & vbCrLf & vbCrLf
        
        If Trim(AdoRs.Fields(8) & "") <> "" Then
            txt_Detail = txt_Detail & "** " & Trim(AdoRs.Fields(8) & "") & vbCrLf
        End If
        txt_Detail = txt_Detail & AdoRs.Fields(9) & vbCrLf & vbCrLf
        
        cmd_Image.Visible = False
        If Val(AdoRs.Fields(10) & "") > 0 Then
            cmd_Image.Visible = True
            iImageCnt = Val(AdoRs.Fields(10) & "")
        End If
        
    End If

    AdoRs.Close
    Set AdoRs = Nothing
    
    For iDx = 1 To 10
        cmd_img_index(iDx).Visible = True
    Next iDx
    
    For iDx = iImageCnt + 1 To 10
        cmd_img_index(iDx).Visible = False
    Next iDx
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Image_Display(iImageNo As Integer)

    Dim iDR         As String
    Dim sFileName   As String
    Dim sFilePath   As String
    Dim sServPath   As String
    Dim iLoc        As Integer
        
    Screen.MousePointer = vbHourglass
    
    Image1.Picture = LoadPicture("")
    With MDIMain.Inet
   
        .Protocol = icFTP
        .URL = sServerIP
        .UserName = sServerID
        .Password = sServerPWD
                    
        sFileName = sPgmId & "_" & iImageNo & ".JPG"
        
        sFilePath = App.Path & "\" & sFileName
        
        If Dir(sFilePath) <> "" Then
            Kill sFilePath
        End If

        iLoc = InStr(1, sServerPATH, "/")
        sServPath = Left(sServerPATH, iLoc) & "Help/" & sFileName
                
        'Server -> Client Copy
        .Execute , "GET " & sServPath & " " & Chr(34) & sFilePath & Chr(34)
        
        Do While .StillExecuting
            DoEvents
        Loop
        .Execute , "quit"
            
        If Dir(sFilePath) <> "" Then
            Image1.Picture = LoadPicture(sFilePath)
            Kill sFilePath
        End If
        
    End With
        
    Screen.MousePointer = vbDefault
End Sub
