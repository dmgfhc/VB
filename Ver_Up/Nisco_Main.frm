VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form Nisco_Main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nisco System"
   ClientHeight    =   1470
   ClientLeft      =   4455
   ClientTop       =   4620
   ClientWidth     =   5850
   Icon            =   "Nisco_Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5850
   Begin Threed.SSPanel SSPanel1 
      Height          =   1365
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2408
      _Version        =   196609
      BackColor       =   14737632
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel lblFileName 
         Height          =   330
         Left            =   90
         Top             =   90
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   582
         Caption         =   ""
         Alignment       =   1
         BackColor       =   14737632
         BackgroundStyle =   1
         BorderEffect    =   0
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin MSComctlLib.ProgressBar PrgDown 
         Height          =   315
         Left            =   135
         Negotiate       =   -1  'True
         TabIndex        =   1
         Top             =   465
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin InDate.ULabel lblState 
         Height          =   465
         Left            =   90
         Top             =   810
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   820
         Caption         =   ""
         Alignment       =   1
         BackColor       =   14737632
         BackgroundStyle =   1
         BorderEffect    =   0
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   225
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
End
Attribute VB_Name = "Nisco_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sServerIP As String            'SERVER IP
Public sServerID As String            'SERVER ID
Public sServerPWD As String           'SERVER PASSWORD
Public sServerPATH As String          'SERVER PATH
Public FILE_SIZE As Double            'FILE SIZE

Private Sub Form_Load()

On Error GoTo Find_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    Set AdoRs = New ADODB.Recordset
    
    Call Gp_FormCenter(Me)
    Me.Show
    
    SSPanel1.Visible = True
    lblFileName.Visible = False
    PrgDown.Visible = False
    lblState.Caption = "欢迎登录南钢板材三级计算机生产管理系统...!!"
    
    If GF_DbConnect = False Then Unload Me
    
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
    
    If sServerIP = "" Or sServerID = "" Or sServerPWD = "" Or sServerPATH = "" Then
        Set AdoRs = Nothing
        Call Gp_MsgBoxDisplay("服务器相关信息不正确...!!")
        Unload Me
        Exit Sub
    End If

    AdoRs.Close
    Set AdoRs = Nothing
    
    Inet.Protocol = icFTP
    Inet.URL = sServerIP
    Inet.UserName = sServerID
    Inet.Password = sServerPWD
    
    'Main.exe FILE SIZE
    FILE_SIZE = Gf_FloatFind(M_CN1, "SELECT SYS_SIZE FROM ZP_VERSION WHERE SUB_SYS = 'M.exe' ")
    Call ExeRun_File("Main.exe", "MAIN SYSTEM")
    
    Unload Me
    Exit Sub
    
Find_Error:

    Set AdoRs = Nothing
    Call Gp_MsgBoxDisplay("系统启动失败...请稍后再试" + Err.Description)
    Unload Me
    
End Sub

Private Sub ExeRun_File(WinId As String, Form_Caption As String)
    
    Dim sQuery As String
    Dim Client_Ver As String
    Dim Server_Ver As String
    Dim sFilePath As String

    Dim lHandle As Long
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    
    lHandle = FindWindow(vbNullString, Form_Caption)
    
    If lHandle <> 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Server Version
    sQuery = "SELECT TRIM(FST_VER) || TRIM(SND_VER) || TRIM(THR_VER) FROM ZP_VERSION WHERE SUB_SYS = 'M.exe' "
    Server_Ver = Gf_CodeFind(M_CN1, sQuery)
    
    'Client Version
    Client_Ver = Trim(Str(GetPrivateProfileInt("Version", WinId, 0, App.Path & "\" & "ENVI.INI")))
    
    Client_Ver = "000000" & Client_Ver
    
    Client_Ver = Right(Client_Ver, 6)
    
    With Inet
    
        If InStr(1, Server_Ver, Client_Ver, vbTextCompare) = 0 Or Dir(App.Path & "\" & WinId) = "" Then
            
            SSPanel1.Visible = True
            lblFileName.Visible = True
            PrgDown.Visible = True
            lblState.Caption = "Main System File Copying....!!"
            
            'Client File Delete
            If Dir(App.Path & "\" & WinId) <> "" Then
                Kill App.Path & "\" & WinId
                lblFileName.Caption = ""
            End If

            PrgDown.Max = FILE_SIZE
            PrgDown.Value = 0
            
            'Server -> Client Copy
            .Execute , "GET " & sServerPATH & WinId & " " & Chr(34) & App.Path & "\" & WinId & Chr(34)
            
            Do While .StillExecuting
            
                Sleep (100)
                
                DoEvents

                lblFileName.Caption = Format(FileLen(App.Path & "\" & WinId) \ 1024, "#,##0") & " KB" & " / " & Format(FILE_SIZE \ 1024, "#,##0") & " KB"

                If FileLen(App.Path & "\" & WinId) > FILE_SIZE Then
                    PrgDown.Value = FILE_SIZE
                Else
                    PrgDown.Value = FileLen(App.Path & "\" & WinId)
                End If

            Loop

            lblFileName.Caption = Format(FileLen(App.Path & "\" & WinId) \ 1024, "#,##0") & " KB" & " / " & Format(FILE_SIZE \ 1024, "#,##0") & " KB"

            If FileLen(App.Path & "\" & WinId) > FILE_SIZE Then
                PrgDown.Value = FILE_SIZE
            Else
                PrgDown.Value = FileLen(App.Path & "\" & WinId)
            End If

            'End
            .Execute , "quit"
            
            Do While .StillExecuting
                DoEvents
            Loop

            Call WritePrivateProfileString("Version", WinId, Server_Ver, App.Path & "\" & "ENVI.INI")

            .Cancel
            Do While .StillExecuting
                DoEvents
            Loop
        
        End If
            
    End With
    
    lHandle = FindWindow(vbNullString, Form_Caption)
    
    If lHandle = 0 Then
        'CALL
        SaveSetting "NISCO", "EXE-FILE", WinId, "1"
        Shell App.Path & "\" & WinId, vbMaximizedFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Err_Handler:

    SSPanel1.Visible = False
    Call Gp_MsgBoxDisplay("系统启动失败...请稍后再试" & Err.Description)
    Screen.MousePointer = vbDefault
    Unload Me
    
End Sub
