VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Form_DIR_SELECT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择保存路径"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   3735
   StartUpPosition =   3  '窗口缺省
   Begin Threed.SSCommand ssc_CANCEL 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   196609
      ForeColor       =   32768
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form_DIR_SELECT.frx":0000
      Caption         =   "取消"
      PictureAlignment=   0
   End
   Begin Threed.SSCommand ssc_OK 
      Height          =   375
      Left            =   420
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   196609
      ForeColor       =   192
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form_DIR_SELECT.frx":0354
      Caption         =   "确认"
      PictureAlignment=   0
   End
   Begin VB.TextBox txt_DIR 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   3960
      Width           =   3315
   End
   Begin VB.DirListBox Dir_SAVE 
      Height          =   3240
      Left            =   180
      TabIndex        =   1
      Top             =   570
      Width           =   3315
   End
   Begin VB.DriveListBox Drive_SAVE 
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   3315
   End
End
Attribute VB_Name = "Form_DIR_SELECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir_SAVE_Change()
    txt_DIR.Text = Trim(Dir_SAVE.Path)
End Sub

Private Sub Drive_SAVE_Change()
    Dir_SAVE.Path = Drive_SAVE.Drive
End Sub

Private Sub Form_Load()
    Dir_SAVE.Path = Drive_SAVE.Drive
End Sub

Private Sub ssc_CANCEL_Click()
    Unload Me
End Sub

Private Sub ssc_OK_Click()
    sEXLSavePATH = Trim(txt_DIR.Text)
    Unload Me
End Sub
