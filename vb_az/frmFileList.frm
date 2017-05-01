VERSION 5.00
Begin VB.Form frmFileList 
   Caption         =   "File Selection"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox otxtFileName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   495
      Width           =   5235
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2715
      TabIndex        =   3
      Top             =   2565
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确认"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1635
      TabIndex        =   2
      Top             =   2565
      Width           =   975
   End
   Begin VB.FileListBox oFile1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   105
      Pattern         =   "*.xls"
      TabIndex        =   1
      Top             =   825
      Width           =   5250
   End
   Begin VB.Label Label1 
      Caption         =   "FilePath := C:\NISCO\EXCEL"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   5115
   End
End
Attribute VB_Name = "frmFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ExcelFilePath  As String
'Const ExcelFilePath = App.Path & "\EXCEL"

Private Sub Form_Activate()

    ExcelFilePath = "C:\NISCO\EXCEL"
    Label1.Caption = "FilePath := " & ExcelFilePath
    otxtFileName = ""
    
    oFile1.Path = ExcelFilePath
    
    oFile1.Pattern = "*.XLS"
    oFile1.SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim vForm As Variant
    
    AZD1010C.lblFileName.Caption = ExcelFilePath & "\" & Trim(otxtFileName)
    Me.Tag = ""

End Sub

Private Sub cmdCancel_Click()
    
    otxtFileName = ""
    Unload Me

End Sub

Private Sub cmdOk_Click()
    
    If otxtFileName > "" Then
        Unload Me
    Else
        MsgBox "File Not Selected"
    End If

End Sub

Private Sub oFile1_Click()
    
    cmdOk.Enabled = True
    otxtFileName = oFile1.FileName

End Sub

Private Sub oFile1_DblClick()
    Call cmdOk_Click
End Sub
