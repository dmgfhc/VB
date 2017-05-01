VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form ACZ1020C 
   BackColor       =   &H8000000A&
   Caption         =   "TEST界面_ACZ1020C"
   ClientHeight    =   7830
   ClientLeft      =   510
   ClientTop       =   2040
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   13920
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand cmdLoad 
      Height          =   405
      Left            =   2460
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   714
      _Version        =   196609
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LOAD"
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   405
      Left            =   555
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   714
      _Version        =   196609
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CLEAR"
   End
   Begin VB.Image Image1 
      Height          =   9315
      Left            =   555
      Top             =   1020
      Width           =   14145
   End
End
Attribute VB_Name = "ACZ1020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Production Management System
'-- Sub_System Name
'-- Program Name      Image Selection
'-- Program ID        ACZ1020C
'-- Designer          KIM SOO HEON
'-- Coder             KIM SOO HEON
'-- Date              2005.12.02
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim SQL             As String
Dim strStream       As adodb.Stream

Private Sub Form_Define()
    Dim I As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")

    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    Dim idr    As Long
    
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
'    Call Gp_Sp_Setting(ss1)
'    Call Gf_Sp_Cls(sc1)
'    Call Gp_Sp_ColGet(ss1, "G-System.INI", Me.Name)
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Screen.MousePointer = vbDefault
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    Call Gp_Sp_ColSet(ss1, "G-System.INI", Me.Name)

    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
''
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1

End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol

End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Image1.Picture = Nothing
    
End Sub

Private Sub cmdLoad_Click()

    HelpDiaplay.Tag = Me.Name
    HelpDiaplay.Show (0)
'
'    Dim idr         As String
'    Dim sFilePath   As String
'    Dim sFileName   As String
'    Dim sServPath   As String
'    Dim iLoc        As Integer
'    Dim sQuery      As String
'    Dim AdoRs       As adodb.Recordset
'
'    On Error Resume Next
'
'    sFileName = Me.Name & ".jpg"
'
'    sFilePath = App.Path & "\" & sFileName
'
'    If Dir(sFilePath) <> "" Then
'        Kill sFilePath
'    End If
'
'    If sServerIP = "" Then
'        Set AdoRs = New adodb.Recordset
'
'        sQuery = "SELECT SERVER_IP, SERVER_ID, SERVER_PWD, SERVER_PATH FROM ZP_SERVERINFO "
'        AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'        If Not AdoRs.BOF And Not AdoRs.EOF Then
'
'            If VarType(AdoRs.Fields(0)) = vbNull Then
'                sServerIP = ""
'            Else
'                sServerIP = AdoRs.Fields(0)
'            End If
'
'            If VarType(AdoRs.Fields(1)) = vbNull Then
'                sServerID = ""
'            Else
'                sServerID = AdoRs.Fields(1)
'            End If
'
'            If VarType(AdoRs.Fields(2)) = vbNull Then
'                sServerPWD = ""
'            Else
'                sServerPWD = AdoRs.Fields(2)
'            End If
'
'            If VarType(AdoRs.Fields(3)) = vbNull Then
'                sServerPATH = ""
'            Else
'                sServerPATH = AdoRs.Fields(3)
'            End If
'
'        End If
'
'        AdoRs.Close
'        Set AdoRs = Nothing
'    End If
'
'    With MDIMain.Inet
'        .Cancel
'
'        .Protocol = icFTP
'        .URL = sServerIP
'        .UserName = sServerID
'        .Password = sServerPWD
'
'        iLoc = InStr(1, sServerPATH, "/")
'        sServPath = Left(sServerPATH, iLoc) & "Help/" & sFileName
'
'        'Server -> Client Copy
'        .Execute , "GET " & sServPath & " " & Chr(34) & sFilePath & Chr(34)
'
'        Do While .StillExecuting
'            DoEvents
'        Loop
'        .Execute , "quit"
'
'    End With
'
'    Image1.Picture = LoadPicture(sFilePath)
'    Kill sFilePath
            
End Sub

'Private Sub cmdSelectSave_Click()
'    Dim oPict       As StdPicture
'    Dim sFileName   As String
'    Dim strTemp     As Boolean
'
'    With File1
'        sFileName = .Path & "\" & .FileName
'        If .FileName = "" Then
'            MsgBox "Invalid filename or file not found.", vbOKOnly + vbExclamation, "错误提示!"
'            Exit Sub
'        End If
'
'        Set oPict = LoadPicture(sFileName)
'        'Exit Function if this is NOT a picture file
'        If oPict Is Nothing Then
'            MsgBox "Invalid Picture File!", vbOKOnly, "错误提示!"
'            Exit Sub
'        End If
'
'        Set strStream = New adoDb.Stream
'        strStream.Type = adTypeBinary
'        strStream.Open
'        strStream.LoadFromFile sFileName
'
'        Set AdoRs = New adoDb.Recordset
'
'        SQL = "SELECT FILE_PATH FROM ZP_HELP_FILE WHERE TRIM(PGMID) =  '" & Trim(.FileName) & "'" & vbCrLf
'
'        AdoRs.Open SQL, adOpenKeyset, adLockOptimistic, adCmdText
'
'        strTemp = FileToBLOB(sFileName, AdoRs!FILE_PATH, False, 8192)
'        If AdoRs.EOF Then
'            SQL = "INSERT INTO ZP_HELP_FILE(FILE_PATH) values("" & strTemp & "")"
'            SQL = SQL & "WHERE TRIM(PGMID) =  '" & Trim(.FileName) & "'" & vbCrLf
'            AdoRs.Update
'        Else
'            SQL = "UPDATE ZP_HELP_FILE SET FILE_PATH = "" & strTemp & """
'            SQL = SQL & " WHERE TRIM(PGMID) =  '" & Trim(.FileName) & "'" & vbCrLf
'            AdoRs.Update
'        End If
'
'        AdoRs.Close
'        Set AdoRs = Nothing
'    End With
'
'End Sub
'
'
'Private Sub File1_Click()
'    Dim oPict       As StdPicture
'    Dim sFileName   As String
'
'    With File1
'
'        If .FileName = "" Then
'            MsgBox "Invalid filename or file not found.", vbOKOnly + vbExclamation, "错误提示!"
'            Exit Sub
'        Else
'            sFileName = .Path & "\" & .FileName
'
'            Set oPict = LoadPicture(sFileName)
'
'            'Exit Function if this is NOT a picture file
'            If oPict Is Nothing Then
'                MsgBox "Invalid Picture File!", vbOKOnly, "错误提示!"
'                Exit Sub
'            End If
'
'            Set strStream = New adoDb.Stream
'            strStream.Type = adTypeBinary
'            strStream.Open
'            strStream.LoadFromFile sFileName
'
'            Image1.Picture = LoadPicture(sFileName)
'        End If
'
'    End With
'End Sub
'
'
'Public Function BLOBToFile(ByVal imgFilePath As String, ByRef objField As adoDb.Field, _
'                               Optional ByVal bUseStream As Boolean = True, _
'                               Optional ByVal lngChunkSize As Long = 8192) As Boolean
'
'    On Error Resume Next
'    Dim objStream As adoDb.Stream
'    Dim intFreeFile As Integer
'    Dim lngBytesLeft As Long
'    Dim lngReadBytes As Long
'    Dim byBuffer() As Byte
'
'    If bUseStream Then
'
'        Set objStream = New adoDb.Stream
'
'        With objStream
'        .Type = adTypeBinary
'        .Open
'        .Write objField.Value
'        .SaveToFile imgFilePath, adSaveCreateOverWrite
'        End With
'
'        DoEvents
'    Else
'
'        If Dir(imgFilePath) <> "" Then
'            Kill imgFilePath
'        End If
'
'        lngBytesLeft = objField.ActualSize
'        intFreeFile = FreeFile
'
'        Open imgFilePath For Binary As #intFreeFile
'
'        Do Until lngBytesLeft <= 0
'            lngReadBytes = lngBytesLeft
'
'            If lngReadBytes > lngChunkSize Then
'                lngReadBytes = lngChunkSize
'            End If
'
'            byBuffer = objField.GetChunk(lngReadBytes)
'            Put #intFreeFile, , byBuffer
'            lngBytesLeft = lngBytesLeft - lngReadBytes
'
'            DoEvents
'        Loop
'
'        Close #intFreeFile
'
'    End If
'
'    If ERR.Number <> 0 Or ERR.LastDllError <> 0 Then
'        BLOBToFile = False
'    Else
'        BLOBToFile = True
'    End If
'
'End Function
'
'Public Function FileToBLOB(ByVal imgFilePath As String, ByRef objField As adoDb.Field, _
'                               Optional ByVal bUseStream As Boolean = True, _
'                               Optional ByVal lngChunkSize As Long = 8192) As Boolean
'    On Error Resume Next
'    Dim objStream As adoDb.Stream
'    Dim intFreeFile As Integer
'    Dim lngBytesLeft As Long
'    Dim lngReadBytes As Long
'    Dim byBuffer() As Byte
'    Dim varChunk As Variant
'
'
'    If bUseStream Then
'        Set objStream = New adoDb.Stream
'        With objStream
'            .Type = adTypeBinary
'            .Open
'            .LoadFromFile imgFilePath
'            objField.Value = .Read(adReadAll)
'        End With
'    Else
'        With objField
'
'            If (.Attributes And adFldLong) <> 0 Then
'                intFreeFile = FreeFile
'                Open imgFilePath For Binary Access Read As #intFreeFile
'                lngBytesLeft = LOF(intFreeFile)
'
'
'                Do Until lngBytesLeft <= 0
'
'                    If lngBytesLeft > lngChunkSize Then
'                        lngReadBytes = lngChunkSize
'                    Else
'                        lngReadBytes = lngBytesLeft
'                    End If
'
'                    ReDim byBuffer(lngReadBytes)
'
'                    Get #intFreeFile, , byBuffer()
'                    objField.AppendChunk byBuffer()
'                    lngBytesLeft = lngBytesLeft - lngReadBytes
'
'                    DoEvents
'                    Loop
'                    Close #intFreeFile
'                Else
'                    ERR.Raise -10000, "FileToBLOB", _
'                    "The Database Field does Not support Long Binary Data."
'                End If
'            End With
'        End If
'
'       If ERR.Number <> 0 Or ERR.LastDllError <> 0 Then
'           FileToBLOB = False
'       Else
'           FileToBLOB = True
'       End If
'End Function

